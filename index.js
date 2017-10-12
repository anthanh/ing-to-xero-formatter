#!/usr/bin/env node

var XLSX = require('xlsx')
const toCSV = require('array-to-csv')
const inputFile = process.argv[2] || './input.xls'
const outputFile = process.argv[3] || './output.csv'
const SHEET_NAME = 'report1'

const COL_FECHA = 'A'
const COL_DESCRIPCIÓN = 'C'
const COL_CONCEPTO = 'D'
const COL_IMPORTE = 'E'
const COL_SALDO = 'F'
const ROW_OFFSET = 6

var workbook = XLSX.readFile(inputFile)
var LENGTH = getWorkbookLength(workbook) - ROW_OFFSET

function getWorkbookLength(workbook) {
  const lastKey = Object.keys(workbook.Sheets[SHEET_NAME]).length - 4
  return +Object.keys(workbook.Sheets[SHEET_NAME])[lastKey].slice(1, Infinity)
}

function getValue(colName, row) {
  const value = workbook.Sheets[SHEET_NAME][colName + (row + ROW_OFFSET)] ?
  workbook.Sheets[SHEET_NAME][colName + (row + ROW_OFFSET)].v : ''
  if (colName === COL_IMPORTE) {
    return value.replace('.', '').replace(',', '.')
  }
  return value
}

// ING Spain export format
// ['FECHA', '', 'DESCRIPCIÓN', 'CONCEPTO', 'IMPORTE', 'SALDO', '']
// Xero CSV format header
// *Date,*Amount,Payee,Description,Reference,Check Number

const data = []
for (let i = 0; i <= LENGTH; i++) {
  var amount = getValue(COL_IMPORTE, i)
  if (isNaN(+amount)) {
    console.log(amount, +amount, i)
  }
  data.push([
    getValue(COL_FECHA, i), +amount,
    null,
    getValue(COL_DESCRIPCIÓN, i),
    getValue(COL_CONCEPTO, i),
    null
  ])
}

var fs = require('fs')
fs.writeFile(outputFile, toCSV(data), function (err) {
  if (err) {
    return console.log(err)
  }
  console.log(`DONE => ${outputFile}`)
})