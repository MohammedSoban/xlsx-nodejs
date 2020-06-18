var xlsx = require('xlsx-style')
var xlsxs = require('xlsx')
var colors = require('colors');
var wb = xlsx.readFile('book1.xlsx', { cellStyles: true })
var ws = wb.Sheets["Sheet1"]

var data = xlsx.utils.sheet_to_json(ws)


console.log(data)

userSum = 33


let sum = 0

let count = 0

data.map((value, index) => {
  // var a=parseFloat(value && value.values.replace(/,/g, ''))
  if (value.values) {
    sum += parseFloat(value && value.values.replace(/,/g, ''))
    count++
  }
})




xlsxs.utils.sheet_add_aoa(ws, [[sum]], { origin: `B${count + 2}` });
xlsxs.utils.sheet_add_aoa(ws, [[userSum]], { origin: `C${count + 2}` });

if (sum === userSum) {




  ws[`B${count + 2}`].s = {
    fill: {
      type: 'pattern',
      pattern: "solid", // none / solid
      fgColor: { rgb: "008000" },
      bgColor: { rgb: "008000" }
    },
  }

  console.log('userSum', userSum)
  console.log('sheet sum', sum)

  console.log("Success!".green)

  xlsx.writeFile(wb, 'Book1.xlsx')

} else {


  ws[`B${count + 2}`].s = {
    fill: {
      type: 'pattern',
      pattern: "solid", // none / solid
      fgColor: { rgb: "FF0000" },
      bgColor: { rgb: "FF0000" }
    },
  }
  console.log('userSum', userSum)
  console.log('sheet sum', sum)
  console.log("Fail!".red)

  xlsx.writeFile(wb, 'Book1.xlsx')
}

count = 0




