const request = require('request');
const Excel = require('excel4node');

const url = 'https://example.com';
const fileName = 'snapshot.xlsx';

request(url, (error, response, body) => {
  if (error) {
    console.log(error);
  } else {
    console.log('Snapshot taken successfully!');
  }
});

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('Sheet 1');

const data = JSON.parse(body);
worksheet.cell(1, 1).string('Name');
worksheet.cell(1, 2).string('Age');
worksheet.cell(1, 3).string('Occupation');
for (let i = 0; i < data.length; i++) {
  worksheet.cell(i + 2, 1).string(data[i].name);
  worksheet.cell(i + 2, 2).number(data[i].age);
  worksheet.cell(i + 2, 3).string(data[i].occupation);
}

workbook.xlsx.writeFile(fileName)
  .then(() => {
    console.log('Snapshot saved successfully!');
  })
  .catch((error) => {
    console.log(error);
  });