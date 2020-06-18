const fs = require('fs')
var colors = require('colors');

var parse = require('csv-parse');
const csv = require('csv-parser')

var excel = require('excel4node');
const readXlsxFile = require('read-excel-file/node');

var workbook = new excel.Workbook();

var worksheet = workbook.addWorksheet('Sheet1');
const csvData = []
const final_result = [];
const userSum = 6.7
// fs.createReadStream('newfile.csv')
//     .pipe(
//         parse({


//             delimiter: ','

//         })
//     ).on('data', function (dataRow) {
//         csvData.push(dataRow)
//         // console.log(dataRow)
//     })
//     .on('end', function () {
//         console.log(csvData)
//         var sum = 0
//         csvData.map((element, index) => {
//             sum = Number(element[0]) + sum
//             // console.log(element[index])

//         })

//         console.log(sum, "file sum")
//         console.log(userSum, "User sum")

//         if (sum === userSum) {
//             console.log("Success!".green)
//         } else {
//             console.log("Fail!".red)
//         }

//     })



// fs.createReadStream("newfile.csv")
// .pipe(csv())
// .on("data", data => final_result.push(data))
// .on("end", () => {
//   console.log(final_result);
// });


// var excel = require('excel4node');

// // Create a new instance of a Workbook class
// var workbook = new excel.Workbook();

// // Add Worksheets to the workbook
// var worksheet = workbook.addWorksheet('Sheet1');


// // Create a reusable style
// var style = workbook.createStyle({
//   font: {
//     color: '#FF0800',
//     size: 12
//   },
//   numberFormat: '$#,##0.00; ($#,##0.00); -'
// });

// const bgStyle = workbook.createStyle({
//   fill: {
//     type: 'pattern',
//     patternType: 'solid',
//     bgColor: '#FF0000',
//     fgColor: '#FF0000',
//   }
// });


// worksheet.cell(2, 2).style(style);

// worksheet.cell(1, 1).style(bgStyle);
// worksheet.cell(1, 1).string('RED');

// // var a = workbook.getExcelCellRef(1, 1);
// // console.log(a)

// // var a =excel.getExcelCellRef(1, 1)
// var a =excel.getExcelRowCol('A2');
// console.log(a)

// workbook.write('file.xlsx');




readXlsxFile('./Book1.xlsx',{ sheet: 'Sheet1' }).then((rows) => {


  const fail = workbook.createStyle({
    fill: {
      type: 'pattern',
      patternType: 'solid',
      bgColor: '#FF0000',
      fgColor: '#FF0000',
    }
  });
  const success = workbook.createStyle({
    fill: {
      type: 'pattern',
      patternType: 'solid',
      bgColor: '#00FF00',
      fgColor: '#00FF00',
    }
  });

  var sum = 0

  rows.map((row, index) => {
    sum = Number(row[0]) + sum
  })


  if (sum === userSum) {
    console.log("Success!".green)
    worksheet.cell(7, 1).style(success);
  } else {
    console.log("Fail!".red)
    worksheet.cell(7, 1).style(fail);
  }

   workbook.write('Book2.xlsx');
}).catch(err => {
  console.log(err)
})



//  readXlsxFile('Book1.xlsx',{ sheet: 'Sheet1' }).then((rows) => {
//   console.log(rows)
// }).catch(err => {
//   console.log(err)
// })