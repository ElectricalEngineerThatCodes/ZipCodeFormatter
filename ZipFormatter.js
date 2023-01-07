'use strict';

//Import Dependencies
const Excel = require('exceljs');
const csv = require('csvtojson');

async function getWorkbook() {
  const donors = await csv().fromFile('UnformattedZipcodes.csv');

  const formattedDonors = getFormattedDonors(donors);

  const {
    nonMetroDetroitMichiganDonors,
    metroDetroitDonors,
    nonMichiganDonors,
  } = populateZipcodes(formattedDonors);

  const workbook = new Excel.Workbook();
  const nonMetroDetroitMichiganDonorsWorksheet =
    workbook.addWorksheet('480XX, 483XX-487XX');
  const metroDetroitDonorsWorksheet =
    workbook.addWorksheet('481XX-482XX, 492XX');
  const nonMichiganDonorsWorksheet = workbook.addWorksheet('Mixed');

  generateExcelSheet(
    nonMetroDetroitMichiganDonorsWorksheet,
    nonMetroDetroitMichiganDonors
  );
  generateExcelSheet(metroDetroitDonorsWorksheet, metroDetroitDonors);
  generateExcelSheet(nonMichiganDonorsWorksheet, nonMichiganDonors);

  // save under export.xlsx
  // await workbook.xlsx.writeFile('FormatForMailingTest.xlsx');

  return workbook;
}

const populateZipcodes = (donors) => {
  const nonMetroDetroitMichiganDonors = [];
  const metroDetroitDonors = [];
  const nonMichiganDonors = [];

  const metroDetroitZipCodePrefixes = new Set([492, 481, 482]);
  const nonMetroDetroitMichiganZipcodesPrefixes = new Set([
    480, 483, 484, 485, 486, 487,
  ]);

  donors.forEach((donor, i) => {
    let first3Zip = Number(donor['Zip Code'].substring(0, 3));

    if (nonMetroDetroitMichiganZipcodesPrefixes.has(first3Zip)) {
      nonMetroDetroitMichiganDonors.push(donor);
    } else if (metroDetroitZipCodePrefixes.has(first3Zip)) {
      metroDetroitDonors.push(donor);
    } else {
      nonMichiganDonors.push(donor);
    }
  });

  return {
    nonMetroDetroitMichiganDonors,
    metroDetroitDonors,
    nonMichiganDonors,
  };
};

const getFormattedDonors = function (donors) {
  return donors.map((donor) => {
    const newDonor = {
      ...donor,
      'Formal Greeting': !donor['Formal Greeting']
        ? donor['Full Name']
        : donor['Formal Greeting'],
      'Informal Greeting': donor['Informal Greeting']
        ? donor['Informal Greeting']
        : donor['First Name'],
      Address: donor['Address 2']
        ? donor['Address'] + ', ' + donor['Address 2']
        : donor['Address'],
    };

    delete newDonor['Address 2'];
    delete newDonor['First Name'];
    delete newDonor['Full Name'];
    delete newDonor['Household Name'];

    return newDonor;
  });
};

const generateExcelSheet = function (worksheet, donors) {
  const columnHeaders = Object.keys(donors[0]);

  columnHeaders.forEach((currHeader, i) => {
    worksheet.getCell(String.fromCharCode(i + 65) + 1).value = currHeader;
  });

  donors.forEach((donor, i) => {
    columnHeaders.forEach((data, j) => {
      let cellAlpha = String.fromCharCode(j + 65);
      let cellNumber = 2 + i;

      worksheet.getCell(cellAlpha + cellNumber).value = donor[data];
    });
  });
};

module.exports = { getWorkbook };
