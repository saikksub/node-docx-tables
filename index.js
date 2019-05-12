'use strict'

const fs = require('fs')
const path = require('path')
const JSZip = require('jszip')
const convert = require('xml-js')

/**
 * Load and Extract given docx file
 */
async function loadFile (file) {
  return new Promise((resolve, reject) => {
    fs.readFile(file, function (err, data) {
      if (err) {
        reject(err)
      }
      JSZip.loadAsync(data).then(function (zip) {
        resolve(zip)
      }).catch((error) => {
        reject(error)
      })
    })
  })
}

/**
 * Main Logic for extracting Table data from XML JSON data
 */
function parseTables (xmlJsonData) {
  const result = {}
  try {
    let wTable = xmlJsonData['w:document']['w:body']['w:tbl']
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
    return error
  }

  return result
}

module.exports = function (props) {
  return new Promise((resolve, reject) => {
    if (!(props && props.constructor === {}.constructor)) {
      reject(new Error(`Invalid Props`))
    }
    if (!props.file) {
      reject(new Error(`Object prop "file" is required.`))
    }
    if (!fs.existsSync(props.file)) {
      reject(new Error(`Input file "${props.file}" does not exists. Please provide valid file.`))
    }
    // Load and extract docx file
    loadFile(props.file).then((data) => {
      if (data.files['word/document.xml']) {
        data.files['word/document.xml'].async("binarystring").then(function (content) {
          // Parse XML data
          let xmlJsonData = convert.xml2json(content, {compact: true, spaces: 4})
          // If the XML data is invalid, resolve empty object
          if (!xmlJsonData) {
            resolve({})
          }
          // Make sure parsed XML file is an object
          if (typeof xmlJsonData === 'string') {
            xmlJsonData = JSON.parse(xmlJsonData)
          }
          
          const result = parseTables(xmlJsonData)
          resolve(result)
        })
      } else {
        resolve({})
      }
    }).catch((error) => {
      reject(error)
    })
  })
}
