// npm install --save docxtemplater pizzip

// PARA IMAGEN ES ESTO : 
// npm install git+https://github.com/pwndoc/docxtemplater-image-module-pwndoc   

// ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ comando para genera el word usar eso

const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require("docxtemplater-image-module-pwndoc");
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
    res.send("hola")
})

app.post('/generar', (req, res) => {

    const {nombre, edad, direccion} = req.body

    const nombretitulo = "vengo_del_backend"

    // Define una ruta relativa al archivo y Lee el contenido binario del archivo de plantilla .docx
    const content = fs.readFileSync(
        path.resolve(__dirname, 'prueba.docx'),
        'binary'
    );

    const imagen1 = "foto.jpg"; // Ruta de la primera imagen a insertar


    // Crea una instancia de PizZip utilizando el contenido del archivo .docx

    const zip = new PizZip(content);
    const imageOptions = {
        centered: false,
        getImage(tagValue, tagName) {
            console.log({ tagValue, tagName });
            return fs.readFileSync(tagValue);
        },
        getSize() {
            // it also is possible to return a size in centimeters, like this : return [ "2cm", "3cm" ];
            return [500, 500];
        },
    };



       // Crea una instancia de Docxtemplater, configurando opciones como 'paragraphLoop' y 'linebreaks'
    const doc = new Docxtemplater(zip, {
        modules: [new ImageModule(imageOptions)],
    });
    

    // Renderiza el documento con los datos proporcionados, osea manda los datos que deseas al word , y esos variables tienes que ponerlo a tu word
    doc.render({
        nombre: nombre,
        edad: edad,
        direccion: direccion,
        'foto': imagen1

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