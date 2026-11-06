require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const express = require('express');
const { Pool } = require('pg');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const archiver = require('archiver');
const ExcelJS = require('exceljs');
const { S3Client, PutObjectCommand, DeleteObjectCommand, GetObjectCommand } = require('@aws-sdk/client-s3');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configurar cliente de Cloudflare R2 (Compatible con S3)
const s3Client = new S3Client({
  region: 'auto',
  endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: {
    accessKeyId: process.env.R2_ACCESS_KEY_ID || '',
    secretAccessKey: process.env.R2_SECRET_ACCESS_KEY || '',
  },
});
const BUCKET_NAME = process.env.R2_BUCKET_NAME;
const R2_PUBLIC_URL = (process.env.R2_PUBLIC_URL || '').replace(/\/$/, '');

// Configurar Multer en memoria (RAM)
const upload = multer({ storage: multer.memoryStorage() });

// Conexión a PostgreSQL Neon (con respaldo directo integrado y sin channel_binding)
const RAW_DB_URL = process.env.DATABASE_URL || 'postgresql://neondb_owner:npg_hQ2pcDf0EnXm@ep-dry-glitter-b4cqe8pc-pooler.c-6.us-east-2.aws.neon.tech/neondb?sslmode=require';
const CLEAN_DB_URL = RAW_DB_URL.replace('&channel_binding=require', '');

const pool = new Pool({
  connectionString: CLEAN_DB_URL,
  ssl: { rejectUnauthorized: false }
});

// Verificar conexión y tablas al iniciar
async function initDB() {
  try {
    const res = await pool.query('SELECT NOW()');
    console.log('✅ Conectado a PostgreSQL (Neon) correctamente:', res.rows[0].now);
  } catch (err) {
    console.error('❌ Error al conectar a la BD:', err.message);
  }
}
initDB();

// Configuración de límites: 10 PDFs, 10 Exceles y 100 Fotografías
const multiUpload = (req, res, next) => {
  const uploader = upload.fields([
    { name: 'pdfs', maxCount: 10 },
    { name: 'excels', maxCount: 10 },
    { name: 'imagenes', maxCount: 100 }
  ]);

  uploader(req, res, function (err) {
    if (err instanceof multer.MulterError) {
      if (err.code === 'LIMIT_UNEXPECTED_FILE') {
        return res.status(400).json({ error: 'Has superado el límite permitido (Máx: 10 PDFs, 10 Exceles o 100 Fotografías).' });
      }
      return res.status(400).json({ error: err.message });
    } else if (err) {
      return res.status(500).json({ error: err.message });
    }
    next();
  });
};

// Función auxiliar para subir archivo a Cloudflare R2
async function subirArchivoAR2(file, carpeta, documentoId) {
  const nombreLimpio = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
  const r2Key = `${carpeta}/${documentoId}_${Date.now()}_${nombreLimpio}`;

  await s3Client.send(new PutObjectCommand({
    Bucket: BUCKET_NAME,
    Key: r2Key,
    Body: file.buffer,
    ContentType: file.mimetype
  }));

  const archivoUrl = R2_PUBLIC_URL ? `${R2_PUBLIC_URL}/${r2Key}` : `/api/r2-file?key=${encodeURIComponent(r2Key)}`;
  return { r2Key, archivoUrl };
}

// ==================== FUNCIONES AUXILIARES ====================

// 1. Obtener partidas filtradas (Pestaña 1 y 2)
async function obtenerPartidasFiltradas(filtros) {
  const { anio, proyecto, valorizacion, partida, estado } = filtros;
  let query = 'SELECT * FROM documentos_reporte WHERE 1=1';
  let params = [];
  let idx = 1;

  if (anio) { query += ` AND CAST(anio AS TEXT) = $${idx++}`; params.push(anio.toString().trim()); }
  if (proyecto) { query += ` AND UPPER(TRIM(proyecto)) = UPPER(TRIM($${idx++}))`; params.push(proyecto); }
  if (valorizacion) { query += ` AND UPPER(TRIM(valorizacion)) = UPPER(TRIM($${idx++}))`; params.push(valorizacion); }
  if (partida) { 
    query += ` AND (partida_codigo ILIKE $${idx} OR partida_descripcion ILIKE $${idx})`; 
    params.push(`%${partida.trim()}%`); 
    idx++; 
  }
  query += ' ORDER BY id ASC';

  const result = await pool.query(query, params);
  const partidas = result.rows;

  if (partidas.length > 0) {
    const ids = partidas.map(p => p.id);
    const archivosRes = await pool.query(
      'SELECT * FROM documento_archivos WHERE documento_id = ANY($1) ORDER BY id ASC',
      [ids]
    );
    const mapaArchivos = {};
    archivosRes.rows.forEach(a => {
      if (!mapaArchivos[a.documento_id]) mapaArchivos[a.documento_id] = [];
      mapaArchivos[a.documento_id].push(a);
    });
    partidas.forEach(p => { p.archivos = mapaArchivos[p.id] || []; });
  }

  if (estado === 'con_archivo') {
    return partidas.filter(p => p.archivos.length > 0);
  } else if (estado === 'sin_archivo') {
    return partidas.filter(p => p.archivos.length === 0);
  }
  return partidas;
}

// 2. Obtener datos de CEPRODER filtrados (Pestaña 3)
async function obtenerDatosCeproder(filtros) {
  const { proyecto, fInicio, fFin, valorizacion, adenda, partidas, comunidad, genero, edad, dni, nombres } = filtros;
  
  let query = 'SELECT * FROM reportes_ceproder WHERE 1=1';
  let values = [];
  let i = 1;

  if (proyecto) {
    query += ` AND UPPER(TRIM(proyecto)) = UPPER(TRIM($${i++}))`;
    values.push(proyecto);
  }
  if (fInicio) {
    query += ` AND fecha >= $${i++}`;
    values.push(fInicio);
  }
  if (fFin) {
    query += ` AND fecha <= $${i++}`;
    values.push(fFin);
  }
  if (valorizacion) {
    query += ` AND (CAST(valorizacion AS TEXT) = $${i} OR CAST(valorizacion AS TEXT) = LPAD($${i}, 2, '0') OR CAST(valorizacion AS TEXT) ILIKE '%' || $${i} || '%')`;
    values.push(valorizacion.toString().trim());
    i++;
  }
  if (adenda) {
    query += ` AND (CAST(adenda AS TEXT) = $${i} OR CAST(adenda AS TEXT) = LPAD($${i}, 2, '0') OR CAST(adenda AS TEXT) ILIKE '%' || $${i} || '%')`;
    values.push(adenda.toString().trim());
    i++;
  }
  if (comunidad) {
    query += ` AND UPPER(TRIM(comunidad)) = UPPER(TRIM($${i++}))`;
    values.push(comunidad);
  }
  if (genero) {
    query += ` AND UPPER(TRIM(genero)) = UPPER(TRIM($${i++}))`;
    values.push(genero);
  }
  if (edad) {
    query += ` AND CAST(edad AS TEXT) = $${i++}`;
    values.push(edad.toString().trim());
  }
  if (dni) {
    query += ` AND CAST(dni AS TEXT) ILIKE $${i++}`;
    values.push(`%${dni.trim()}%`);
  }
  if (nombres) {
    query += ` AND nombres_apellidos ILIKE $${i++}`;
    values.push(`%${nombres.trim()}%`);
  }
  if (partidas && partidas.trim().length > 0) {
    const arrayPartidas = partidas.split(',').map(p => p.trim()).filter(Boolean);
    if (arrayPartidas.length > 0) {
      query += ` AND TRIM(CAST(cod_partida AS TEXT)) = ANY($${i++})`;
      values.push(arrayPartidas);
    }
  }

  query += ' ORDER BY comunidad ASC, nombres_apellidos ASC LIMIT 5000';

  const result = await pool.query(query, values);
  return result.rows;
}

// ==================== RUTAS PESTAÑA 1 Y 2 ====================

app.get('/api/documentos', async (req, res) => {
  try {
    const finalResult = await obtenerPartidasFiltradas(req.query);
    res.json(finalResult);
  } catch (err) {
    console.error('Error en GET /api/documentos:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/documentos', multiUpload, async (req, res) => {
  try {
    const { anio, proyecto, valorizacion, partida_codigo, partida_descripcion } = req.body;
    const insertRes = await pool.query(
      `INSERT INTO documentos_reporte (anio, proyecto, valorizacion, partida_codigo, partida_descripcion)
       VALUES ($1, $2, $3, $4, $5) RETURNING id`,
      [anio, proyecto, valorizacion, partida_codigo, partida_descripcion]
    );
    const docId = insertRes.rows[0].id;

    const grupos = [
      { campo: 'pdfs', categoria: 'pdf' },
      { campo: 'excels', categoria: 'excel' },
      { campo: 'imagenes', categoria: 'imagen' }
    ];

    for (const g of grupos) {
      const files = (req.files && req.files[g.campo]) || [];
      for (const file of files) {
        const { r2Key, archivoUrl } = await subirArchivoAR2(file, g.categoria, docId);
        await pool.query(
          `INSERT INTO documento_archivos (documento_id, categoria, nombre_original, archivo_url, r2_key)
           VALUES ($1, $2, $3, $4, $5)`,
          [docId, g.categoria, file.originalname, archivoUrl, r2Key]
        );
      }
    }

    res.json({ ok: true, id: docId });
  } catch (err) {
    console.error('Error en POST /api/documentos:', err);
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/documentos/:id', multiUpload, async (req, res) => {
  try {
    const docId = req.params.id;
    const { anio, proyecto, valorizacion, partida_codigo, partida_descripcion } = req.body;

    await pool.query(
      `UPDATE documentos_reporte
       SET anio = $1, proyecto = $2, valorizacion = $3, partida_codigo = $4, partida_descripcion = $5
       WHERE id = $6`,
      [anio, proyecto, valorizacion, partida_codigo, partida_descripcion, docId]
    );

    const grupos = [
      { campo: 'pdfs', categoria: 'pdf' },
      { campo: 'excels', categoria: 'excel' },
      { campo: 'imagenes', categoria: 'imagen' }
    ];

    for (const g of grupos) {
      const files = (req.files && req.files[g.campo]) || [];
      for (const file of files) {
        const { r2Key, archivoUrl } = await subirArchivoAR2(file, g.categoria, docId);
        await pool.query(
          `INSERT INTO documento_archivos (documento_id, categoria, nombre_original, archivo_url, r2_key)
           VALUES ($1, $2, $3, $4, $5)`,
          [docId, g.categoria, file.originalname, archivoUrl, r2Key]
        );
      }
    }

    res.json({ ok: true });
  } catch (err) {
    console.error('Error en PUT /api/documentos/:id:', err);
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/archivos/:id', async (req, res) => {
  try {
    const fileRes = await pool.query('SELECT * FROM documento_archivos WHERE id = $1', [req.params.id]);
    if (fileRes.rows.length > 0) {
      const archivo = fileRes.rows[0];
      if (archivo.r2_key) {
        await s3Client.send(new DeleteObjectCommand({ Bucket: BUCKET_NAME, Key: archivo.r2_key })).catch(() => {});
      }
      await pool.query('DELETE FROM documento_archivos WHERE id = $1', [req.params.id]);
    }
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/documentos/:id', async (req, res) => {
  try {
    const archivosRes = await pool.query('SELECT * FROM documento_archivos WHERE documento_id = $1', [req.params.id]);
    for (const arch of archivosRes.rows) {
      if (arch.r2_key) {
        await s3Client.send(new DeleteObjectCommand({ Bucket: BUCKET_NAME, Key: arch.r2_key })).catch(() => {});
      }
    }
    await pool.query('DELETE FROM documentos_reporte WHERE id = $1', [req.params.id]);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/documentos/:id/zip', async (req, res) => {
  try {
    const docRes = await pool.query('SELECT * FROM documentos_reporte WHERE id = $1', [req.params.id]);
    if (docRes.rows.length === 0) return res.status(404).send('Partida no encontrada');
    const doc = docRes.rows[0];

    const archivosRes = await pool.query('SELECT * FROM documento_archivos WHERE documento_id = $1', [req.params.id]);
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="Partida_${doc.partida_codigo}.zip"`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    for (const arch of archivosRes.rows) {
      if (arch.r2_key) {
        const obj = await s3Client.send(new GetObjectCommand({ Bucket: BUCKET_NAME, Key: arch.r2_key }));
        archive.append(obj.Body, { name: `${arch.categoria.toUpperCase()}/${arch.nombre_original}` });
      }
    }

    await archive.finalize();
  } catch (err) {
    console.error('Error al generar ZIP:', err);
    res.status(500).send('Error al generar ZIP: ' + err.message);
  }
});

app.get('/api/r2-file', async (req, res) => {
  try {
    const key = req.query.key;
    if (!key) return res.status(400).send('Falta el parámetro key');
    const obj = await s3Client.send(new GetObjectCommand({ Bucket: BUCKET_NAME, Key: key }));
    if (obj.ContentType) res.setHeader('Content-Type', obj.ContentType);
    obj.Body.pipe(res);
  } catch (err) {
    res.status(404).send('Archivo no encontrado en R2');
  }
});

app.get('/api/reporte-excel', async (req, res) => {
  try {
    const partidas = await obtenerPartidasFiltradas(req.query);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Sistema de Gestión Documental';
    workbook.created = new Date();

    const wsDetalle = workbook.addWorksheet('1. Detalle de Partidas');
    wsDetalle.mergeCells('A1:J1');
    const titleCell = wsDetalle.getCell('A1');
    titleCell.value = 'REPORTE DETALLADO DE SUSTENTOS Y ARCHIVOS POR PARTIDA';
    titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    wsDetalle.getRow(1).height = 28;

    const headersDetalle = [
      'N°', 'Año', 'Proyecto', 'Valorización', 'Código', 'Descripción de la Partida',
      'Estado', 'Archivos PDF Cargados', 'Archivos Excel Cargados', 'Fotografías Cargadas'
    ];
    const headerRow = wsDetalle.addRow(headersDetalle);
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    let totalConArchivos = 0;
    let totalSinArchivos = 0;
    let sumaPdfs = 0;
    let sumaExcels = 0;
    let sumaFotos = 0;

    partidas.forEach((p, index) => {
      const pdfs = p.archivos.filter(a => a.categoria === 'pdf').map(a => `• ${a.nombre_original}`);
      const excels = p.archivos.filter(a => a.categoria === 'excel').map(a => `• ${a.nombre_original}`);
      const fotos = p.archivos.filter(a => a.categoria === 'imagen').map(a => `• ${a.nombre_original}`);

      sumaPdfs += pdfs.length;
      sumaExcels += excels.length;
      sumaFotos += fotos.length;

      const tieneArchivos = p.archivos.length > 0;
      if (tieneArchivos) totalConArchivos++; else totalSinArchivos++;

      const row = wsDetalle.addRow([
        index + 1, p.anio, p.proyecto, p.valorizacion, p.partida_codigo, p.partida_descripcion,
        tieneArchivos ? 'COMPLETO / CON ARCHIVOS' : 'PENDIENTE',
        pdfs.length > 0 ? pdfs.join('\n') : 'Sin archivos',
        excels.length > 0 ? excels.join('\n') : 'Sin archivos',
        fotos.length > 0 ? fotos.join('\n') : 'Sin archivos'
      ]);

      row.eachCell((cell, colNumber) => {
        cell.alignment = {
          vertical: 'top',
          horizontal: [1, 2, 4, 5, 7].includes(colNumber) ? 'center' : 'left',
          wrapText: true
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
        if (colNumber === 7) {
          cell.font = { bold: true, color: { argb: tieneArchivos ? 'FF15803D' : 'FFB45309' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tieneArchivos ? 'FFDCFCE7' : 'FFFEF3C7' } };
        }
      });
    });

    wsDetalle.columns = [
      { width: 6 }, { width: 8 }, { width: 38 }, { width: 20 }, { width: 12 },
      { width: 42 }, { width: 22 }, { width: 32 }, { width: 32 }, { width: 32 }
    ];

    const wsResumen = workbook.addWorksheet('2. Resumen');
    wsResumen.mergeCells('A1:C1');
    const resTitle = wsResumen.getCell('A1');
    resTitle.value = 'RESUMEN EJECUTIVO DE DOCUMENTACIÓN';
    resTitle.font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    resTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    resTitle.alignment = { vertical: 'middle', horizontal: 'center' };
    wsResumen.getRow(1).height = 26;

    const addTablaResumen = (titulo, filas, startRow) => {
      const hRow = wsResumen.getRow(startRow);
      hRow.values = [titulo, 'Cantidad', 'Porcentaje / Detalle'];
      hRow.height = 20;
      hRow.eachCell(c => {
        c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      filas.forEach(f => {
        const r = wsResumen.addRow(f);
        r.eachCell((c, colIdx) => {
          c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          c.alignment = { vertical: 'middle', horizontal: colIdx === 1 ? 'left' : 'center' };
          if (colIdx === 1) c.font = { bold: true };
        });
      });
    };

    const totalPartidas = partidas.length || 1;
    const porcCon = ((totalConArchivos / totalPartidas) * 100).toFixed(1) + '%';
    const porcSin = ((totalSinArchivos / totalPartidas) * 100).toFixed(1) + '%';

    addTablaResumen('Indicador de Partidas (Según Filtro)', [
      ['Total de Partidas Evaluadas', partidas.length, '100%'],
      ['Partidas con Sustento Cargado', totalConArchivos, porcCon],
      ['Partidas Pendientes de Sustento', totalSinArchivos, porcSin]
    ], 3);

    addTablaResumen('Desglose de Archivos Cargados por Tipo', [
      ['Total de Documentos PDF', sumaPdfs, 'Archivos .pdf'],
      ['Total de Hojas de Cálculo Excel', sumaExcels, 'Archivos .xls / .xlsx'],
      ['Total de Fotografías / Imágenes', sumaFotos, 'Imágenes'],
      ['TOTAL GENERAL DE ARCHIVOS', sumaPdfs + sumaExcels + sumaFotos, 'Todos los formatos']
    ], 8);

    wsResumen.columns = [{ width: 38 }, { width: 18 }, { width: 25 }];

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Reporte_Valorizaciones_Partidas.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).send('Error al generar reporte Excel: ' + err.message);
  }
});

// ==================== RUTAS PESTAÑA 3 (CEPRODER) ====================

// 1. Obtener datos con filtros
app.get('/api/ceproder', async (req, res) => {
  try {
    const data = await obtenerDatosCeproder(req.query);
    res.json(data);
  } catch (error) {
    console.error('❌ Error en GET /api/ceproder:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// 2. Exportar a Excel con filtros aplicados
app.get('/api/ceproder/exportar', async (req, res) => {
  try {
    const data = await obtenerDatosCeproder(req.query);
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Registros CEPRODER');

    worksheet.columns = [
      { header: 'N°', key: 'nro_registro', width: 10 },
      { header: 'NOMBRES Y APELLIDOS', key: 'nombres_apellidos', width: 35 },
      { header: 'DNI', key: 'dni', width: 15 },
      { header: 'EDAD', key: 'edad', width: 10 },
      { header: 'GÉNERO', key: 'genero', width: 10 },
      { header: 'COMUNIDAD', key: 'comunidad', width: 20 },
      { header: 'SECTOR', key: 'sector', width: 20 },
      { header: 'CARGO', key: 'cargo', width: 20 },
      { header: 'ACTIVIDAD / DESCRIPCION', key: 'actividad', width: 40 },
      { header: 'PROYECTO', key: 'proyecto', width: 25 },
      { header: 'VALORIZACION', key: 'valorizacion', width: 15 },
      { header: 'COD. PARTIDA', key: 'cod_partida', width: 15 }
    ];

    data.forEach(row => worksheet.addRow(row));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Reporte_CEPRODER.xlsx"');
    
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).send(error.message);
  }
});

// 3. Obtener Partidas únicas para los filtros
app.get('/api/ceproder/partidas', async (req, res) => {
  try {
    const result = await pool.query('SELECT DISTINCT cod_partida FROM reportes_ceproder WHERE cod_partida IS NOT NULL ORDER BY cod_partida');
    res.json(result.rows.map(r => r.cod_partida));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ==================== ENCENDER SERVIDOR ====================
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`🚀 Servidor corriendo en http://localhost:${PORT}/reportes.html`));