'use strict'

const fs = require('fs')
const JSZip = require('jszip')
const convert = require('xml-js')
const fsExtra = require('fs-extra')

fs.readFile('./test2.docx', function (err, data) {
  if (err) throw err
  JSZip.loadAsync(data).then(function (zip) {
    if (zip.files['word/document.xml']) {
      zip.files['word/document.xml'].async("binarystring").then(function (content) {
        let xml = convert.xml2json(content, {compact: true, spaces: 4})
        const result = {}
        if (typeof xml === 'string') {
          xml = JSON.parse(xml)
        }
        fsExtra.writeJsonSync('./doc.json', xml, { spaces: 4 })
        try {
          let wTable = xml['w:document']['w:body']['w:tbl']
          if (wTable) {
            if (wTable.constructor !== [].constructor) {
              wTable = [ wTable ]
            }
            wTable.forEach((wTableItem) => {
              const wTableItemRow = wTableItem['w:tr']
              wTableItemRow.forEach((wTableItemRowItem, rowIndex) => {
                const wTableItemRowColumn = wTableItemRowItem['w:tc']
                const rowObject = []
                wTableItemRowColumn.forEach((wTableItemRowColumnItem, colIndex) => {
                  let wp = wTableItemRowColumnItem['w:p']
                  if (wp) {
                    if ( wp.constructor !== [].constructor) {
                      wp = [ wp ]
                    }
                    let data = ''
                    wp.forEach((wpItem) => {
                      if (wpItem['w:r'] && wpItem['w:r']['w:t'] && wpItem['w:r']['w:t']._text) {
                        // console.log(wpItem['w:r']['w:t'])
                        data += `${wpItem['w:r']['w:t']._text}\n`
                      }
                    })
                    if (data) {
                      rowObject.push({
                        position: {
                          row: rowIndex,
                          col: colIndex
                        },
                        data 
                      })
                    }
                  }
                  // console.log('++++++++++++++++++')
                })
                if (rowObject && rowObject.constructor === [].constructor && rowObject.length > 0) {
                  result[`${rowIndex}`] = Object.assign([], rowObject)
                }
                // console.log('==========================')
              }) 
            })
          }
        } catch (error) {
          throw error
        }

        fsExtra.writeJsonSync('./out.json', result, { spaces: 4 })
      })
    }
  })
})
