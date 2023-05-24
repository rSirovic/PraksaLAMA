const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

app.post('/generate-excel', async (req, res) => {
  try {
    const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
    await createExcelFile(data);
    res.send('Excel file created successfully.');
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('An error occurred while generating the Excel file.');
  }
});


async function createExcelFile(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');

  const logoImage = workbook.addImage({
    filename: 'logo.jpg', 
    extension: 'jpg', 
  })


  worksheet.addImage(logoImage, {
    tl: { col: 0.5, row: 1 }, 
    br: { col: 2, row: 3.25 },
  });

 
  worksheet.mergeCells('A5:C5');
  const subject = data.profesori[0];
  const mergedCell = worksheet.getCell('A5');
  mergedCell.value = subject ? 'Predmet: ' + subject['PredmetNaziv'] + ' (' + subject['PredmetKratica'] + ')' : 'Predmet: N/A';

  worksheet.mergeCells('A6:I11');
  worksheet.getCell('A6').value =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.';
  worksheet.getCell('A6').alignment = { wrapText: true };

  // Outside boarder
  const columnLetterss = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
  const rowNumber = 13;
  
  columnLetterss.forEach((letter) => {
    const cell = worksheet.getCell(letter + rowNumber);
    const borderStyle = { bottom: { style: 'medium' }, right: { style: 'thin' } };
    if (letter === 'I') {
      borderStyle.right.style = 'medium';
    }
    cell.border = borderStyle;
  });


  // celije za profesore
  const mergeCellsConfig = [
    'A15:A16',
    'B15:B16',
    'C15:C16',
    'D15:D16',
    'E15:G15',
    'H15:H16',
    'I15:I16',
    'J15:J16',
    'K15:M15',
    'N15:N16'
  ];
  
  mergeCellsConfig.forEach((mergeRange) => {
    worksheet.mergeCells(mergeRange);
  });

  const cellValuesConfig = [
    { cell: 'A15', value: 'Redni broj' },
    { cell: 'B15', value: 'Ime i Prezime' },
    { cell: 'C15', value: 'Zvanje' },
    { cell: 'D15', value: 'Status' },
    { cell: 'E15', value: 'Sati Nastave' },
    { cell: 'E16', value: 'pred' },
    { cell: 'F16', value: 'sem' },
    { cell: 'G16', value: 'vjez' },
    { cell: 'H15', value: 'Bruto satnica predavanja (EUR)' },
    { cell: 'I15', value: 'Bruto satnica seminari (EUR)' },
    { cell: 'J15', value: 'Bruto satnica vjezbe (EUR)' },
    { cell: 'K15', value: 'Bruto iznos' },
    { cell: 'K16', value: 'pred' },
    { cell: 'L16', value: 'sem' },
    { cell: 'M16', value: 'vjez' },
    { cell: 'N15', value: 'Ukupno za isplatu (EUR)' }
  ];
  
  cellValuesConfig.forEach(({ cell, value }) => {
    worksheet.getCell(cell).value = value;
  });

    
    worksheet.mergeCells('A12:B12');
    worksheet.mergeCells('H12:I12');
  
    
    const cellValues = [
      'Katedra',
      null, // Preskakanje prazne ćelije
      'Studij',
      'ak. god.',
      'stud. god.',
      'pocetak turnusa',
      'kraj turnusa',
      'broj sati predviden programom'
    ];
    
    const columnLetters = ['A', 'C', 'D', 'E', 'F', 'G', 'H'];
    
    columnLetters.forEach((letter, index) => {
      const cell = worksheet.getCell(letter + '12');
      cell.value = cellValues[index] || ''; // Ako vrijednost u polju nije definirana, koristimo prazan string
    });
    
  
    
    const katedra = data.profesori;
    const row = worksheet.getRow(13);
    
    const katedraSatiPred = katedra.reduce((total, professor) => total + professor.PlaniraniSatiPredavanja, 0);
    const katedraSatiSem = katedra.reduce((total, professor) => total + professor.PlaniraniSatiSeminari, 0);
    const katedraSatiVjez = katedra.reduce((total, professor) => total + professor.PlaniraniSatiVjezbe, 0);
  
    worksheet.mergeCells(`A13:B13`);
    worksheet.getCell('A13').value = katedra[0]['Katedra'];
    worksheet.getCell('C13').value = katedra[0]['Studij'];
    worksheet.getCell('D13').value = katedra[0]['SkolskaGodinaNaziv'];
    worksheet.getCell('E13').value = katedra[0]['PKSkolskaGodina'];
    worksheet.getCell('F13').value = katedra[0]['PocetakTurnusa'];
    worksheet.getCell('G13').value = katedra[0]['KrajTurnusa'];
    worksheet.mergeCells('H13:I13');
    worksheet.getCell('H13').value =
      'P: ' + 
      katedraSatiPred + ' ' +
      'S: ' + 
      katedraSatiSem + ' ' +
      'V: ' +
      + katedraSatiVjez;
  
    row.alignment = { horizontal: 'left' };
    worksheet.getCell('H13').alignment = { horizontal: 'center' };
  

  // Formatiranje headera
  const headerRows = [15, 16, 12];
const rowStylesConfig = [
  {
    font: { bold: true },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E7E7E7' } },
    border: { top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' } }
  }
];

headerRows.forEach((rowNumber) => {
  const row = worksheet.getRow(rowNumber);

  rowStylesConfig.forEach((styleConfig) => {
    Object.keys(styleConfig).forEach((property) => {
      row[property] = styleConfig[property];
    });
  });

  row.eachCell((cell) => {
    cell.fill = rowStylesConfig[0].fill;
    cell.border = rowStylesConfig[0].border;
  });
});

  // Visina headera
  const cellHeights = {
    A12: 50,
    D16: 100,
    E16: 100,
    F16: 100,
  };

  Object.entries(cellHeights).forEach(([cellRef, height]) => {
    const cell = worksheet.getCell(cellRef);
    const row = worksheet.getRow(cell.row);
    row.height = height;
  });

  
  data.profesori.forEach((professor, index) => {
    const rowNumber = 16 + index + 1;
    worksheet.getCell(`A${rowNumber}`).value = index + 1;
    worksheet.getCell(`B${rowNumber}`).value = professor['NastavnikSuradnikNaziv'];
    worksheet.getCell(`C${rowNumber}`).value = professor['Zvanje'];
    worksheet.getCell(`D${rowNumber}`).value = professor['NazivNastavnikStatus'];
    worksheet.getCell(`E${rowNumber}`).value = professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`F${rowNumber}`).value = professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`G${rowNumber}`).value = professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`H${rowNumber}`).value = professor['NormaPlaniraniSatiPredavanja'];
    worksheet.getCell(`I${rowNumber}`).value = professor['NormaPlaniraniSatiSeminari'];
    worksheet.getCell(`J${rowNumber}`).value = professor['NormaPlaniraniSatiVjezbe'];
    worksheet.getCell(`K${rowNumber}`).value = professor['NormaPlaniraniSatiPredavanja'] * professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`L${rowNumber}`).value = professor['NormaPlaniraniSatiSeminari'] * professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`M${rowNumber}`).value = professor['NormaPlaniraniSatiVjezbe'] * professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`N${rowNumber}`).value = worksheet.getCell(`K${rowNumber}`).value + worksheet.getCell(`L${rowNumber}`).value + worksheet.getCell(`M${rowNumber}`).value;

    const sum = worksheet.getCell(`K${rowNumber}`).value + worksheet.getCell(`L${rowNumber}`).value + worksheet.getCell(`M${rowNumber}`).value;
    worksheet.getCell(`N${rowNumber}`).value = sum ;
    worksheet.getCell(`N${rowNumber}`).border = { right: {style: 'medium'}} ;

  
    worksheet.getRow(rowNumber).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
  });

  // Sirine stupaca A do J

  worksheet.columns[0].width = 6;
  worksheet.columns[1].width = 18.43;
  worksheet.columns[2].width = 21.14;
  worksheet.columns[3].width = 21.14;
  worksheet.columns[4].width = 6.14;
  worksheet.columns[5].width = 7.86;
  worksheet.columns[6].width = 8.14;
  worksheet.columns[7].width = 10.14;
  worksheet.columns[8].width = 10;
  worksheet.columns[9].width = 10.14;
  worksheet.columns.forEach((column, index) => {
    if (index >= 10) {
      column.width = 8.43;
    }
  });

  const totalRowNumber = 16 + data.profesori.length + 1;

  worksheet.mergeCells(`A${totalRowNumber}:C${totalRowNumber}`);
  worksheet.getCell(`A${totalRowNumber}`).value = 'Ukupno';
  worksheet.getCell(`A${totalRowNumber}`).alignment = { horizontal: 'center' };

  // Racunanje sati
  const ccolumnLetters = ['E', 'F', 'G'];
  const ttotalRowNumber = 16 + data.profesori.length + 2;
  
  ccolumnLetters.forEach((ccolumnLetter) => {
    worksheet.getCell(`${ccolumnLetter}${ttotalRowNumber}`).value = {
      formula: `SUM(${ccolumnLetter}17:${ccolumnLetter}${ttotalRowNumber - 1})`,
    };
  });
  

  //Bruto satnica
  const cLetters = ['H', 'I', 'J'];
  const totalRowNumber1 = 16 + data.profesori.length + 2;
  
  cLetters.forEach((cLetter) => {
    worksheet.getCell(`${cLetter}${totalRowNumber1}`).value = {
      formula: `SUM(${cLetter}17:${cLetter}${totalRowNumber1 - 1})`,
    };
  });
  

  // Bruto iznosi
  const columnLetters1 = ['K', 'L', 'M'];
  const totalRowNumber2 = 16 + data.profesori.length + 2;
  
  columnLetters1.forEach((columnLetter1) => {
    worksheet.getCell(`${columnLetter1}${totalRowNumber2}`).value = {
      formula: `SUM(${columnLetter1}17:${columnLetter1}${totalRowNumber2 - 1})`,
    };
  });
  

  //Ukupan iznos
  worksheet.getCell(`N${totalRowNumber}`).value = {
    formula: `SUM(K${totalRowNumber}:M${totalRowNumber})`
  };


  const dekani = data.dekani;

  const dekanRows = [
    { row: totalRowNumber + 3, index: 0, position: 'Prodekanica za nastavu i studentska pitanja' },
    { row: totalRowNumber + 9, index: 1, position: 'Prodekan za financije i upravljanje' },
    { row: totalRowNumber + 9, index: 2, position: 'Dekan' }
  ];
  
  dekanRows.forEach(({ row, index, position }) => {
    const cellAddress = index === 2 ? `J${row}` : `A${row}`;
    const dekan = dekani[index];
  
    worksheet.mergeCells(`${cellAddress}:${index === 2 ? 'L' : 'C'}${row + 1}`);
    worksheet.getCell(cellAddress).value = {
      richText: [
        { text: `${position}\n` },
        { text: `Prof. dr. sc. ${dekan.ImePrezime}` },
      ],
    };
    worksheet.getCell(cellAddress).alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
  });
  

    
    await workbook.xlsx.writeFile('output.xlsx');
    console.log('Excel file created successfully.');
  }

app.listen(3000, () => {
  console.log('API server je pokrenut na portu 3000.');
});
  
  
