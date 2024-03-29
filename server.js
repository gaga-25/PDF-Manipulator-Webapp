const express = require('express')
const app = express()
const path = require('path')
const multer= require('multer')
const fs = require('fs').promises;

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
const {mergedPDFs} = require('./merge')
const { PDFNet } = require('@pdftron/pdfnet-node');

const upload1  = multer({dest: 'uploads1/'})
const upload2 = multer({dest: 'uploads2/'})
const upload3 = multer({dest: 'uploads3/'})
const upload4 = multer({dest: 'uploads4/'})
const upload5 = multer({dest: 'uploads5/'})


app.use('/static', express.static('files'))
app.use('/static', express.static('files2'))
app.use('/static', express.static('files3'))
app.use('/static', express.static('files4'))
app.use('/static', express.static('files5'))

app.use(express.static(__dirname +'/css'));

const port= 3000
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, "home.html"))
})
app.get('/merger', (req,res) =>{
  res.sendFile(path.join(__dirname, "merger.html"))
})
app.post('/merge', upload1.array('pdfs',2), async(req, res, next) =>{
let d = await mergedPDFs(path.join(__dirname, req.files[0].path), path.join(__dirname, req.files[1].path))
res.redirect(`http://localhost:3000/static/${d}.pdf`)
})
app.get('/convertToPdf' , (req, res) => {
  res.sendFile(path.join(__dirname, "convertToPdf.html"))
})
app.post('/convertToPdf', upload2.single('file'), async(req, res, next) =>{
  const ext = '.pdf'
  const filename = req.file.originalname.substring(0,req.file.originalname.lastIndexOf("."))
  const inputPath =  path.join(__dirname, req.file.path); 
  const outputPath = path.resolve(__dirname, `./files2/${filename}.pdf`);
  const docxBuf = await fs.readFile(inputPath);
  let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
  await fs.writeFile(outputPath, pdfBuf)
  res.redirect(`http://localhost:3000/static/${filename}.pdf`)
})
app.get('/convertPdfToDoc' , (req, res) => {
  res.sendFile(path.join(__dirname, "convertPdfToDoc.html"))
})
app.post('/convertPdfToDoc', upload3.single('file'), async(req, res, next) =>{
  async function main() {
    const filename = req.file.originalname.substring(0,req.file.originalname.lastIndexOf("."))
    const inputPath = path.join(__dirname, req.file.path);
    const outputPath = path.resolve(__dirname, `./files3/${filename}.docx`);
    await PDFNet.addResourceSearchPath('./StructuredOutputWindows/Lib/Windows/');
    if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
      return;
    }
    await PDFNet.Convert.fileToWord(inputPath, outputPath);
    res.redirect(`http://localhost:3000/static/${filename}.docx`)
  }
  PDFNet.runWithCleanup(main, '000b281730c4022043cb420475b561095cddf229b1a');
  
})


app.get('/convertPdfToExcel' , (req, res) => {
  res.sendFile(path.join(__dirname, "convertPdfToExcel.html"))
})

app.post('/convertPdfToExcel', upload4.single('file'), async(req, res, next) =>{
  async function main() {
    const filename = req.file.originalname.substring(0,req.file.originalname.lastIndexOf("."))
    const inputPath = path.join(__dirname, req.file.path);
    const outputPath = path.resolve(__dirname, `./files4/${filename}.xlsx`);
    await PDFNet.addResourceSearchPath('./StructuredOutputWindows/Lib/Windows/');
    if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
      return;
    }
    await PDFNet.Convert.fileToExcel(inputPath, outputPath);
    res.redirect(`http://localhost:3000/static/${filename}.xlsx`)
  }
  PDFNet.runWithCleanup(main, '000b281730c4022043cb420475b561095cddf229b1a');
  
})



app.get('/convertPdfToPpt' , (req, res) => {
  res.sendFile(path.join(__dirname, "convertPdfToPpt.html"))
})

app.post('/convertPdfToPpt', upload5.single('file'), async(req, res, next) =>{
  async function main() {
    const filename = req.file.originalname.substring(0,req.file.originalname.lastIndexOf("."))
    const inputPath = path.join(__dirname, req.file.path);
    const outputPath = path.resolve(__dirname, `./files5/${filename}.pptx`);
    await PDFNet.addResourceSearchPath('./StructuredOutputWindows/Lib/Windows/');
    if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
      return;
    }
    await PDFNet.Convert.fileToPowerPoint(inputPath, outputPath);
    res.redirect(`http://localhost:3000/static/${filename}.pptx`)
  }
  PDFNet.runWithCleanup(main, '000b281730c4022043cb420475b561095cddf229b1a');
  
})



app.listen(port, () => {
  console.log(`Example app listening on port http://localhost:${port}`)
})
