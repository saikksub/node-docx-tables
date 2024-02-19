'use strict'

const fs = require('fs')
const JSZip = require('jszip')
const xml2js = require('xml2js');


/**
 * Load and Extract given docx file
 */
async function loadFile(file) {
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
function parseTables(xmlJsonData) {
  const tables = []
  try {
    let wTable = xmlJsonData['w:document']['w:body']['w:tbl']
    if (wTable) {
      if (wTable.constructor !== [].constructor) {
        wTable = [wTable]
      }
      wTable.forEach((wTableItem) => {
        const result = {}
        const wTableItemRow = wTableItem['w:tr']
        wTableItemRow.forEach((wTableItemRowItem, rowIndex) => {
          const wTableItemRowColumn = wTableItemRowItem['w:tc']
          const rowObject = []
          wTableItemRowColumn.forEach((wTableItemRowColumnItem, colIndex) => {
            let wp = wTableItemRowColumnItem['w:p']
            if (wp) {
              if (wp.constructor !== [].constructor) {
                wp = [wp]
              }
              let data = ''
              wp.forEach((wpItem) => {
                if (wpItem['w:r'] && wpItem['w:r']['w:t'] && (typeof wpItem['w:r']['w:t'] === 'string' || typeof wpItem['w:r']['w:t']._text === 'string')) {
                  data += `${Buffer.from(wpItem['w:r']['w:t'] || wpItem['w:r']['w:t']._text, 'binary').toString('utf-8')}\n`
                }
              })
              //if (data) {
              rowObject.push({
                position: {
                  row: rowIndex,
                  col: colIndex
                },
                data
              })
              //}
            }
            // console.log('++++++++++++++++++')
          })
          //if (rowObject && rowObject.constructor === [].constructor && rowObject.length > 0) {
          result[`${rowIndex}`] = Object.assign([], rowObject)
          //}
          // console.log('==========================')
        })
        tables.push(result)
      })
    }
  } catch (error) {
    return error
  }

  return tables
}

// Function to remove Byte Order Mark (BOM) if it exists
function removeBOM(content) {
  if (content.charCodeAt(0) === 0xFEFF) {
    return content.slice(1);
  }
  return content;
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
      const documentKey = Object.keys(data.files).find(key => /word\/document.*\.xml/.test(key))

      if (data.files[documentKey]) {
        data.files[documentKey].async("binarystring").then(function (content) {
          if (content.startsWith('ï»¿')) {
            content = content.substring(3);
          }

          const parser = new xml2js.Parser({ explicitArray: false, ignoreAttrs: false })
          parser.parseString(content, (err, result) => {
            if (err) {
              throw err;
            }

            let xmlJsonData = JSON.stringify(result, null, 4);

            // Make sure parsed XML file is an object
            if (typeof xmlJsonData === 'string') {
              xmlJsonData = JSON.parse(xmlJsonData)
            }

            const res = parseTables(xmlJsonData)
            resolve(res)
          })
        })
      } else {
        resolve({})
      }
    }).catch((error) => {
      reject(error)
    })
  })
}
