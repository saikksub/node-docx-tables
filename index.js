'use strict'

const fs = require('fs')
const JSZip = require('jszip')

fs.readFile('./test.docx', function (err, data) {
  if (err) throw err
  JSZip.loadAsync(data).then(function (zip) {
    if (zip.files['word/document.xml']) {
      zip.files['word/document.xml'].async("binarystring").then(function (content) {
        console.log(content.split("<w:tbl>").length)
      })
    }
  })
})
