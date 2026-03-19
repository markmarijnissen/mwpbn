import dayjs from "dayjs";

export async function addPlanToDashboard(file, wb, index) {
  const planData = file.data;

  // --- TAB 1: Contactgegevens ---
  const sheet1 = wb.worksheets.find(s => s.name.includes('1. Contactgegevens'));
  if (sheet1) {
    const rIdx = index + 1; // Row 2 for index 1, Row 3 for index 2, etc.
    const row = sheet1.getRow(rIdx);

    // In exceljs, columns are 1-based (1 = A, 2 = B, etc.)
    row.getCell(1).value = index;
    row.getCell(2).value = planData.contactgegevens.collectief;
    row.getCell(3).value = planData.contactgegevens.contactpersoon?.naam || '';
    row.getCell(4).value = planData.contactgegevens.contactpersoon?.functie || '';
    row.getCell(5).value = planData.contactgegevens.contactpersoon?.email || '';
    row.getCell(6).value = planData.contactgegevens.contactpersoon?.telefoon || '';
    row.getCell(7).value = planData.contactgegevens.procesbegeleider?.naam || '';
    row.getCell(8).value = planData.contactgegevens.procesbegeleider?.email || '';
    row.getCell(9).value = planData.contactgegevens.procesbegeleider?.telefoon || '';

    // Flat map betrokkenen into the subsequent columns
    let colOffset = 10;
    planData.contactgegevens.betrokkenen.forEach(b => {
      row.getCell(colOffset++).value = b.naam;
      row.getCell(colOffset++).value = b.functie;
    });

    row.commit(); // Save the row changes
  }

  // --- Helper for TAB 2 & 4 ---
  const fillMatrixSheet = (sheetName, dataObj) => {
    const sheet = wb.worksheets.find(s => s.name.includes(sheetName));
    if (!sheet) return;

    // remove old
    sheet.conditionalFormattings = sheet.conditionalFormattings.filter(cf => cf.ref !== "H4:AU45");
    // re-add
    sheet.addConditionalFormatting({
      ref: "H4:AU45",
      rules: [
        {
          type: "cellIs", operator: "equal", formulae: [0],
          style: { font: { color: { argb: "FFBBBBBB"} } , fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "FFEEEEEE"} } },
        },
        {
          type: "cellIs", operator: "equal", formulae: [1],
          style: { font: { color: { argb: "FF9C0006"} }, fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "FFFFC7CE"} } },
        },
        {
          type: "cellIs", operator: "equal", formulae: [2],
          style: { font: { color: { argb: "FF006100"} }, fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "ffC6EFCE"} } },
        },
        {
          type: "cellIs", operator: "equal", formulae: [3],
          style: { font: { color: { argb: "FF9C5700"} }, fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "FFFFEB9C" } } },
        },
      ]

    })

    // Columns: exceljs is 1-based. 
    const nameCol = index + 7;
    const startValCol = index + 7;
    const toelichtingOffset = 41;
    const startToelCol = startValCol + toelichtingOffset;

    // Write Collectief Name in Row 2
    sheet.getRow(2).getCell(nameCol).value = planData.contactgegevens.collectief;
    sheet.getRow(2).getCell(nameCol + toelichtingOffset).value = planData.contactgegevens.collectief;

    // Write the data matrix (Rows 3 to 42)
    for (let nr = 1; nr <= 40; nr++) {
      const rIdx = nr + 3;
      if (dataObj[nr]) {
        sheet.getRow(rIdx).getCell(startValCol).value = dataObj[nr].waarde;
        // if(dataObj[nr].toelichting) {
        //   sheet.getRow(rIdx).getCell(startValCol).note = dataObj[nr].toelichting;
        // }
        sheet.getRow(rIdx).getCell(startToelCol).value = dataObj[nr].toelichting;
      }
    }
  };

  fillMatrixSheet('2. Opgaven', planData.opgaven);
  fillMatrixSheet('4. (provinciale) Samenwerking', planData.samenwerking);

  // --- TAB 3: Aanbod ---
  const sheet3 = wb.worksheets.find(s => s.name.includes('3. Aanbod'));
  if (sheet3) {
    const colIdx = index + 4; // Assuming it starts at column E (5) for index 1

    
    sheet3.conditionalFormattings = sheet3.conditionalFormattings.filter(cf => cf.ref !== "E4:AR137");
    sheet3.addConditionalFormatting({
      ref: "E4:AR137",
      rules: [
        {
          type: "cellIs", operator: "equal", formulae: [0],
          style: { font: { color: { argb: "FFBBBBBB"} } , fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "FFEEEEEE"} } },
        },
        {
          type: "cellIs", operator: "equal", formulae: [1],
          style: { font: { color: { argb: "FF006100"} }, fill: {type: 'pattern', pattern: 'solid',  bgColor: { argb: "ffC6EFCE"} } },
        }
      ]
    })

    // Write Collectief name in Row 1
    sheet3.getRow(2).getCell(colIdx).value = planData.contactgegevens.collectief;

    // Iterate through rows to find matching "Aanbod" text in Column 1 (A)
    sheet3.eachRow((row, rowNumber) => {
      if (rowNumber > 2) { // Skip headers
        const cellValue = row.getCell(1).value;
        const aanbodText = cellValue ? cellValue.toString().trim() : '';

        if (aanbodText) {
          const isJa = planData.aanbod[aanbodText];
          if (isJa !== undefined) {
            row.getCell(colIdx).value = isJa ? 1 : 0;
          }
        }
      }
    });
  }

  // --- TAB 5: Maatwerkplan collectief ---
  const sheet5 = wb.worksheets.find(s => s.name.includes('5. Maatwerkplan'));
  if (sheet5) {
    // 1. Bepaal de juiste kolom. Index 1 = Kolom D (4). Dus targetCol = 3 + index.
    const targetCol = index + 3;

    // Zet voor de zekerheid de naam van het collectief in rij 2 (zoals in de andere tabs)
    sheet5.getRow(2).getCell(targetCol).value = planData.contactgegevens.collectief;

    // 2. Scan kolom C (3) om de ROW start numbers 
    const prefixes = {
      aanpak: 0,
      voortgang: 0,
      status: 0,
      planning: 0,
      trekker: 0
    };

    // Loop door alle cellen in kolom C
    sheet5.getColumn(3).eachCell((cell, rowNumber) => {
      const cellText = String(cell.value || '').toLowerCase();

      // Check of de cel één van onze attributen bevat (case insensitive)
      for (const key of Object.keys(prefixes)) {
        if (cellText.includes(key)) {
          // Gevonden! Bewaar het rijnummer als de prefix
          prefixes[key] = rowNumber;
        }
      }
    });

    // 3. Vul de data in voor elk criterium (nr)
    if (planData.maatwerkplan) {
      for (const nr in planData.maatwerkplan) {
        const item = planData.maatwerkplan[nr];
        const numericNr = parseInt(nr);

        if (isNaN(numericNr)) continue; // Sla over als het geen geldig nummer is

        // Loop door de attributen in data.maatwerkplan (aanpak, status, etc.)
        for (const key of Object.keys(item)) {
          // Als we een geldige startrij (prefix) hebben gevonden in het dashboard voor deze key
          if (prefixes[key] && prefixes[key] > 0) {

            // Doelrij = ATTRIBUTE_PREFIX + nr (bijv. 3 + 1 = 4)
            const targetRow = prefixes[key] + numericNr;

            // Schrijf de waarde naar de cel!
            const cell = sheet5.getRow(targetRow).getCell(targetCol);
            cell.value = item[key];
          }
        }
      }
    }
  }

  const sheetBestanden = wb.worksheets.find(s => s.name.toLowerCase().includes('bestand'));
  if (sheetBestanden) {
    const row = 1 + index;
    sheetBestanden.getRow(row).getCell(1).value = file.filename;
    sheetBestanden.getRow(row).getCell(2).value = planData.contactgegevens.collectief;
    sheetBestanden.getRow(row).getCell(3).value = dayjs(planData.modified).format("YYYY-MM-DD HH:mm:ss");
    for (let i = 0; i < 6; i++) {
      sheetBestanden.getRow(row).getCell(4 + i).value = file.data.status[i];
    }
    sheetBestanden.getRow(row).getCell(10).value = file.errors.join("\n");
  }
}