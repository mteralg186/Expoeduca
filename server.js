const express = require('express');
const app = express();
const PORT = 3000;
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Middleware para servir archivos estáticos
app.use(express.static('public'));

// Middleware para parsear datos del formulario
app.use(express.urlencoded({ extended: true }));

// Ruta principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
  });

// Ruta para procesar el formulario (POST)
app.post('/contact', (req, res) => {
  const { name, email, inquiry, message } = req.body;
  const filePath = path.join(__dirname, 'contactos.xlsx');

  let workbook;
  let worksheet;

  // Si el archivo ya existe, lo carga
  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);
    data.push({ Nombre: name, Email: email, Asunto: inquiry, Mensaje: message });
    const newWorksheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
  } else {
    // Si no existe, lo crea con una hoja nueva
    const data = [{ Nombre: name, Email: email, Asunto: inquiry, Mensaje: message }];
    worksheet = XLSX.utils.json_to_sheet(data);
    workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contactos');
  }

  // Guardar el archivo
  XLSX.writeFile(workbook, filePath);

  // Enviar respuesta al usuario
  res.send(`<h1>¡Gracias por contactarnos, ${name}!</h1><p>Tu mensaje ha sido guardado correctamente.</p>`);
});

// Arrancar el servidor
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});