const express = require('express');
const { Pool } = require('pg');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const archiver = require('archiver');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Asegurar que exista la carpeta public/uploads
const uploadDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

// Configurar almacenamiento con Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + Math.round(Math.random() * 1E9) + path.extname(file.originalname))
});
const upload = multer({ storage });

// Conexión a PostgreSQL Neon
const pool = new Pool({
  connectionString: process.env.DATABASE_URL || 'postgresql://neondb_owner:npg_hQ2pcDf0EnXm@ep-dry-glitter-b4cqe8pc-pooler.c-6.us-east-2.aws.neon.tech/neondb?sslmode=require&channel_binding=require',
  ssl: { rejectUnauthorized: false }
});

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

// Función auxiliar para obtener datos filtrados (usada tanto en la tabla web como en el Excel)
async function obtenerPartidasFiltradas(filtros) {
  const { anio, proyecto, valorizacion, partida, estado } = filtros;
  let query = 'SELECT * FROM documentos_reporte WHERE 1=1';
  let params = [];
  let idx = 1;

  if (anio) { query += ` AND anio = $${idx++}`; params.push(anio); }
  if (proyecto) { query += ` AND proyecto = $${idx++}`; params.push(proyecto); }
  if (valorizacion) { query += ` AND valorizacion = $${idx++}`; params.push(valorizacion); }
  if (partida) { 
    query += ` AND (partida_codigo ILIKE $${idx} OR partida_descripcion ILIKE $${idx})`; 
    params.push(`%${partida}%`); 
    idx++; 
  }
  query += ' ORDER BY id ASC';

  const result = await pool.query(query, params);
  const partidas = result.rows;

  for (let p of partidas) {
    const archivosRes = await pool.query('SELECT * FROM documento_archivos WHERE documento_id = $1 ORDER BY id ASC', [p.id]);
    p.archivos = archivosRes.rows;
  }

  if (estado === 'con_archivo') {
    return partidas.filter(p => p.archivos.length > 0);
  } else if (estado === 'sin_archivo') {
    return partidas.filter(p => p.archivos.length === 0);
  }
  return partidas;
}

// 1. Obtener todas las partidas con sus archivos
app.get('/api/documentos', async (req, res) => {
  try {
    const finalResult = await obtenerPartidasFiltradas(req.query);
    res.json(finalResult);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// 2. Generar Reporte Excel con 2 Hojas (Detalle y Resumen) según filtros
app.get('/api/reporte-excel', async (req, res) => {
  try {
    const partidas = await obtenerPartidasFiltradas(req.query);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Sistema de Gestión Documental';
    workbook.created = new Date();

    // ==================== HOJA 1: DETALLE DE PARTIDAS ====================
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
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
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
        index + 1,
        p.anio,
        p.proyecto,
        p.valorizacion,
        p.partida_codigo,
        p.partida_descripcion,
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

    // ==================== HOJA 2: RESUMEN EJECUTIVO ====================
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

// 3. Descargar ZIP seguro usando archivo temporal y res.download
app.get('/api/documentos/:id/zip', async (req, res) => {
  try {
    const { id } = req.params;
    const docRes = await pool.query('SELECT * FROM documentos_reporte WHERE id = $1', [id]);
    if (docRes.rows.length === 0) return res.status(404).send('Registro no encontrado');

    const doc = docRes.rows[0];
    const filesRes = await pool.query('SELECT * FROM documento_archivos WHERE documento_id = $1', [id]);
    const archivos = filesRes.rows;

    if (archivos.length === 0) {
      return res.status(400).send('No hay archivos adjuntos en esta partida.');
    }

    const baseLimpia = `${doc.partida_codigo} ${doc.partida_descripcion}`
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[<>:"/\\|?*]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 80);

    const tempZipPath = path.join(uploadDir, `temp_${id}_${Date.now()}.zip`);
    const output = fs.createWriteStream(tempZipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      res.download(tempZipPath, `${baseLimpia}.zip`, () => {
        if (fs.existsSync(tempZipPath)) fs.unlinkSync(tempZipPath);
      });
    });

    archive.on('error', (err) => {
      if (fs.existsSync(tempZipPath)) fs.unlinkSync(tempZipPath);
      if (!res.headersSent) res.status(500).send('Error al generar ZIP: ' + err.message);
    });

    archive.pipe(output);

    let agregados = 0;
    for (let f of archivos) {
      const filePath = path.join(__dirname, 'public', f.archivo_url);
      if (fs.existsSync(filePath)) {
        let subcarpeta = `${baseLimpia}_Otros`;
        if (f.categoria === 'pdf') subcarpeta = `${baseLimpia}_PDFs`;
        else if (f.categoria === 'excel') subcarpeta = `${baseLimpia}_Excels`;
        else if (f.categoria === 'imagen') subcarpeta = `${baseLimpia}_Fotos`;

        const nombreSeguro = path.basename(f.nombre_original).replace(/[<>:"/\\|?*]/g, '_');
        archive.file(filePath, { name: `${subcarpeta}/${nombreSeguro}` });
        agregados++;
      }
    }

    if (agregados === 0) {
      archive.abort();
      if (fs.existsSync(tempZipPath)) fs.unlinkSync(tempZipPath);
      return res.status(404).send('Los archivos físicos no se encontraron en el servidor.');
    }

    await archive.finalize();
  } catch (err) {
    if (!res.headersSent) res.status(500).send('Error interno: ' + err.message);
  }
});

// 4. Crear partida y subir archivos
app.post('/api/documentos', multiUpload, async (req, res) => {
  const { anio, proyecto, valorizacion, partida_codigo, partida_descripcion } = req.body;
  try {
    const docRes = await pool.query(
      `INSERT INTO documentos_reporte (anio, proyecto, valorizacion, partida_codigo, partida_descripcion) 
       VALUES ($1, $2, $3, $4, $5) RETURNING id`,
      [anio, proyecto, valorizacion, partida_codigo, partida_descripcion]
    );
    const documento_id = docRes.rows[0].id;

    const procesarArchivos = async (files, categoria) => {
      if (!files) return;
      for (let f of files) {
        await pool.query(
          `INSERT INTO documento_archivos (documento_id, categoria, nombre_original, archivo_url, tipo_mimetype, tamanio_kb) 
           VALUES ($1, $2, $3, $4, $5, $6)`,
          [documento_id, categoria, f.originalname, `uploads/${f.filename}`, f.mimetype, Math.round(f.size / 1024)]
        );
      }
    };

    if (req.files) {
      await procesarArchivos(req.files['pdfs'], 'pdf');
      await procesarArchivos(req.files['excels'], 'excel');
      await procesarArchivos(req.files['imagenes'], 'imagen');
    }

    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// 5. Actualizar partida o agregar más archivos
app.put('/api/documentos/:id', multiUpload, async (req, res) => {
  const { id } = req.params;
  const { anio, proyecto, valorizacion, partida_codigo, partida_descripcion } = req.body;
  try {
    const existentes = await pool.query(
      'SELECT categoria, COUNT(*) as total FROM documento_archivos WHERE documento_id = $1 GROUP BY categoria',
      [id]
    );
    const conteo = { pdf: 0, excel: 0, imagen: 0 };
    existentes.rows.forEach(r => conteo[r.categoria] = parseInt(r.total));

    const nuevosPdfs = (req.files && req.files['pdfs']) ? req.files['pdfs'].length : 0;
    const nuevosExcels = (req.files && req.files['excels']) ? req.files['excels'].length : 0;
    const nuevasImgs = (req.files && req.files['imagenes']) ? req.files['imagenes'].length : 0;

    if (conteo.pdf + nuevosPdfs > 10) return res.status(400).json({ error: `Esta partida ya tiene ${conteo.pdf} PDF(s). Máximo permitido: 10.` });
    if (conteo.excel + nuevosExcels > 10) return res.status(400).json({ error: `Esta partida ya tiene ${conteo.excel} Excel(es). Máximo permitido: 10.` });
    if (conteo.imagen + nuevasImgs > 100) return res.status(400).json({ error: `Esta partida ya tiene ${conteo.imagen} Fotografía(s). Máximo permitido: 100.` });

    await pool.query(
      `UPDATE documentos_reporte SET anio=$1, proyecto=$2, valorizacion=$3, partida_codigo=$4, partida_descripcion=$5 WHERE id=$6`,
      [anio, proyecto, valorizacion, partida_codigo, partida_descripcion, id]
    );

    const procesarArchivos = async (files, categoria) => {
      if (!files) return;
      for (let f of files) {
        await pool.query(
          `INSERT INTO documento_archivos (documento_id, categoria, nombre_original, archivo_url, tipo_mimetype, tamanio_kb) 
           VALUES ($1, $2, $3, $4, $5, $6)`,
          [id, categoria, f.originalname, `uploads/${f.filename}`, f.mimetype, Math.round(f.size / 1024)]
        );
      }
    };

    if (req.files) {
      await procesarArchivos(req.files['pdfs'], 'pdf');
      await procesarArchivos(req.files['excels'], 'excel');
      await procesarArchivos(req.files['imagenes'], 'imagen');
    }

    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// 6. Eliminar un archivo individual
app.delete('/api/archivos/:id', async (req, res) => {
  try {
    const fileRes = await pool.query('SELECT archivo_url FROM documento_archivos WHERE id = $1', [req.params.id]);
    if (fileRes.rows.length > 0) {
      const filePath = path.join(__dirname, 'public', fileRes.rows[0].archivo_url);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      await pool.query('DELETE FROM documento_archivos WHERE id = $1', [req.params.id]);
    }
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// 7. Eliminar partida completa y sus archivos
app.delete('/api/documentos/:id', async (req, res) => {
  try {
    const filesRes = await pool.query('SELECT archivo_url FROM documento_archivos WHERE documento_id = $1', [req.params.id]);
    for (let f of filesRes.rows) {
      const filePath = path.join(__dirname, 'public', f.archivo_url);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    }
    await pool.query('DELETE FROM documentos_reporte WHERE id = $1', [req.params.id]);
    res.json({ success: true });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Servidor corriendo en el puerto ${PORT}`));