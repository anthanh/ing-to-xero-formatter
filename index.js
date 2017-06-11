#!/usr/bin/env node

const xlsParse = require('xls-parse');
const toCSV = require('array-to-csv');
const inputFile = process.argv[2] || './input.xls';
const outputFile = process.argv[3] || './output.csv';
const sheetName = 'report1';
let data = xlsParse.formatXls2Obj(inputFile, sheetName);

data = data[sheetName];
// remove header
data.splice(0, 6);

// ING Spain export format
// ['FECHA', '', 'DESCRIPCIÓN', 'CONCEPTO', 'IMPORTE', 'SALDO', '']
// Original statement 
// amount is splited in 2 cells due number format
// ['09/06/2017', '', 'Transferencia to XXX', 'INV-0001', '"-2.640', '83"', '"3.988', '25"', '']

data = data.map((statement) => {
  var amount = statement[4].replace(/"|\./g, '') + '.' + statement[5].replace(/"|\./g, '');
  return [statement[0], +amount, null, statement[2], statement[3], null];
});

// Xero CSV format header
// *Date,*Amount,Payee,Description,Reference,Check Number
// console.log(toCSV(data));
var fs = require('fs');
fs.writeFile(outputFile, toCSV(data), function (err) {
  if (err) {
    return console.log(err);
  }
  console.log(`DONE => ${outputFile}`);
});