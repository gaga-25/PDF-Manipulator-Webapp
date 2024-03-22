const PDFMerger = require('pdf-merger-js');

var merger = new PDFMerger();

const mergedPDFs = (async (p1, p2) => {
  await merger.add(p1);  
  await merger.add(p2); // merge only page 2
   //merge pages 3 to 5 (3,4,5)
let d = new Date().getTime()
  await merger.save(`files/${d}.pdf`); //save under given name and reset the internal document
  return d
  
})
module.exports={mergedPDFs}
