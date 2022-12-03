/*
Author: Andrew Townley
Date: 11/29/2022
Description: Format CSV Files to Categorize Alike Zip-Codes
*/
'use strict';

//Import Dependencies
const Excel = require('exceljs');
const csv = require('csvtojson');

(async function () {
  //Load in CSV to Array
  const data = await csv().fromFile('UnformattedZipcodes.csv');

  //Makes Array for Different Zipcode Groups
  const group1 = [];
  const group2 = [];
  const group3 = [];

  //Loop Through Entire Array
  for (let i = 0; i < data.length; i++) {
    let first3Zip = Number(data[i]['Zip Code'].substring(0, 3));

    //First Group of Zip Codes
    if (first3Zip === 480 || (first3Zip <= 487 && first3Zip >= 483)) {
      group1.push(data[i]);
    } else if (first3Zip === 492 || first3Zip === 481 || first3Zip === 482) {
      group2.push(data[i]);
    } else {
      group3.push(data[i]);
    }
  }

  //Deletes Unwanted Data from CSV File
  const deleteData = function (arr) {
    for (let i = 0; i < arr.length; i++) {
      //No Formal Means Use Full Name
      if (!arr[i]['Formal Greeting']) {
        arr[i]['Formal Greeting'] = arr[i]['Full Name'];
      }

      if (!arr[i]['Informal Greeting']) {
        arr[i]['Informal Greeting'] = arr[i]['First Name'];
      }

      // Combine Address 1 && 2
      if (arr[i]['Address 2']) {
        arr[i]['Address'] = arr[i]['Address'] + ', ' + arr[i]['Address 2'];
      }

      //Delete Extras
      delete arr[i]['Address 2'];
      delete arr[i]['First Name'];
      delete arr[i]['Full Name'];
      delete arr[i]['Household Name'];
    }

    return Object.keys(arr[0]);
  };

  const group1Keys = deleteData(group1);
  const group2Keys = deleteData(group2);
  const group3Keys = deleteData(group3);

  //Create New Workbook
  const workbook = new Excel.Workbook();

  // Naming of Wanted Worksheets
  const worksheet1 = workbook.addWorksheet('480XX, 483XX-487XX');
  const worksheet2 = workbook.addWorksheet('481XX-482XX, 492XX');
  const worksheet3 = workbook.addWorksheet('Mixed');

  // Generate excel worksheet Function
  const genExclSht = function (worksheet, columnTitles, arr) {
    //Makes Column Headings
    for (let i = 0; i < columnTitles.length; i++) {
      worksheet.getCell(String.fromCharCode(i + 65) + 1).value =
        columnTitles[i];
    }

    //Enters Data
    for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < columnTitles.length; j++) {
        worksheet.getCell(String.fromCharCode(j + 65) + (2 + i)).value =
          arr[i][`${columnTitles[j]}`];
      }
    }
  };

  //Make the Worksheets for Each Group
  genExclSht(worksheet1, group1Keys, group1);
  genExclSht(worksheet2, group2Keys, group2);
  genExclSht(worksheet3, group3Keys, group3);

  // save under export.xlsx
  await workbook.xlsx.writeFile('FormatForMailing.xlsx');
})();
