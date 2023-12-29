// npm install --save docxtemplater pizzip

// ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ comando para genera el word usar eso

const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const port = 4000;

// Aplica el middleware CORS a todas las rutas
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Middleware para permitir encabezados personalizados
app.use((req, res, next) => {
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    next();
  });

app.get('/', (req, res) =>{
    res.send("hola soy marco")
})

app.post('/generar', (req, res) => {

    const {nombre, edad, direccion} = req.body

    const nombretitulo = "vengo_del_backend"

    // Define una ruta relativa al archivo y Lee el contenido binario del archivo de plantilla .docx
    const content = fs.readFileSync(
        path.resolve(__dirname, 'prueba.docx'),
        'binary'
    );
    // Crea una instancia de PizZip utilizando el contenido del archivo .docx

    const zip = new PizZip(content);

       // Crea una instancia de Docxtemplater, configurando opciones como 'paragraphLoop' y 'linebreaks'
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Renderiza el documento con los datos proporcionados, osea manda los datos que deseas al word , y esos variables tienes que ponerlo a tu word
    doc.render({
        nombre: nombre,
        edad: edad,
        direccion: direccion,

    });

    // Genera el documento como un 'nodebuffer'
    const documentoword = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
    });

 // Configura los encabezados de la respuesta
res.setHeader('Content-Type', 'application/msword'); // Tipo MIME para documentos Word
res.setHeader('Content-Disposition', `${nombretitulo}`); // Nombre del archivo para la descarga

// Envía el documento generado y el título como respuesta al cliente
res.send(documentoword);


});

app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});
