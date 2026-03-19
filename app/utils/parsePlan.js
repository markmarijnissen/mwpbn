import ExcelJS from 'exceljs';

export async function parsePlan(item) {
  item.errors = [];
  const missingSheets = [];
  
  // Maak een nieuw ExcelJS werkboek en laad de ArrayBuffer asynchroon
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(item.content);

  // Helper om een tabblad te vinden op (gedeeltelijke) naam en om te zetten naar een 2D array
  const getSheetData = (partialName) => {
    const sheet = wb.worksheets.find(s => s.name.includes(partialName));
    if (!sheet) {
      missingSheets.push(`Tabblad "${partialName}" niet gevonden`);
      return [];
    }
    
    const rows = [];
    sheet.eachRow({ includeEmpty: true }, (row) => {
      const rowData = [];
      // includeEmpty zorgt ervoor dat we lege cellen niet overslaan, wat de indexering zou breken
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let val = cell.value;
        
        // ExcelJS geeft formules of rich text soms als een object terug. We halen de platte tekst/waarde eruit.
        if (val && typeof val === 'object') {
          if (val.richText) {
            val = val.richText.map(rt => rt.text).join('');
          } else if (val.result !== undefined) {
            val = val.result;
          }
        }
        
        // colNumber is 1-based, we maken er een 0-based array van voor compatibiliteit met je bestaande code
        rowData[colNumber - 1] = val;
      });
      rows.push(rowData);
    });
    
    return rows;
  };

  const data = {
    modified: wb.modified,
    contactgegevens: {
      collectief: '',
      contactpersoon: {},
      procesbegeleider: {},
      betrokkenen: []
    },
    opgaven: {},     // Tab 2
    aanbod: {},      // Tab 3
    samenwerking: {}, // Tab 4
    maatwerkplan: {},
    status: "❌❌❌❌ ".split(''),
    valid: true
  };

  // --- TAB 1: Contactgegevens ---
  const tab1 = getSheetData('1. Contactgegevens');
  let currentSection = '';
  
  for (let r = 0; r < tab1.length; r++) {
    const row = tab1[r];
    const firstCell = String(row[1] || '').trim();
    
    if (firstCell === 'Naam collectief') {
      data.contactgegevens.collectief = row[2] || '';
    } else if (firstCell === 'Contactpersoon doorontwikkeling vanuit collectief') {
      currentSection = 'contactpersoon';
    } else if (firstCell === 'Betrokkenen opstellen maatwerkplan vanuit collectief') {
      currentSection = 'betrokkenen';
    } else if (firstCell === 'Contactpersoon doorontwikkeling (procesbegeleider) vanuit landelijke programmateam') {
      currentSection = 'procesbegeleider';
    } else if (firstCell === 'Naam' && currentSection !== 'betrokkenen') {
      data.contactgegevens[currentSection].naam = row[2] || '';
    } else if (firstCell === 'Functie' && currentSection !== 'betrokkenen') {
      data.contactgegevens[currentSection].functie = row[2] || '';
    } else if (firstCell === 'E-mail' && currentSection !== 'betrokkenen') {
      data.contactgegevens[currentSection].email = row[2] || '';
    } else if (firstCell === 'Telefoonnummer' && currentSection !== 'betrokkenen') {
      data.contactgegevens[currentSection].telefoon = row[2] || '';
    } else if (currentSection === 'betrokkenen' && firstCell && firstCell !== 'Naam') {
      data.contactgegevens.betrokkenen.push({ naam: firstCell, functie: row[3] || '' });
    }
  }
  if(!data.contactgegevens.collectief) item.errors.push(`Contactgegevens niet ingevuld.`);
  if(data.contactgegevens.betrokkenen.length === 0) item.errors.push(`Niemand betrokken.`);
  data.status[0] = data.contactgegevens.betrokkenen.length;

  // --- TAB 2: Opgaven doorontwikkeling ---
  const tab2 = getSheetData('2. Opgaven doorontwikkeling');
  
  let n = 0, count = 0;
  for (let r = 0; r < tab2.length; r++) {
    const row = tab2[r];
    const nr = parseInt(row[1]);
    if (!isNaN(nr)) {
      const voldoen = row[3] === true || String(row[3]).toLowerCase() === 'true';
      const ontwikkel = row[4] === true || String(row[4]).toLowerCase() === 'true';
      
      let val = 0;
      if (voldoen && ontwikkel) val = 3;
      else if (voldoen) val = 1;
      else if (ontwikkel) val = 2;

      data.opgaven[nr] = {
        waarde: val,
        toelichting: row[5] || ''
      };
      if(val >= 2) n++;
      if(val >= 0) count++;
    }
  }
  if(count === 0 || n === 0) item.errors.push(`Ontwikkelpunten niet ingevuld`);
  data.status[1] = n;

  // --- TAB 3: Aanbod ---
  const tab3 = getSheetData('3. Aanbod');
  let currentNr = null;
  n = 0;
  count = 0;
  for (let r = 0; r < tab3.length; r++) {
    const row = tab3[r];
    // Nr is in col 1, Aanbod is in col 4, "Ja" is in col 5
    if (row[1]) currentNr = parseInt(row[1]);
    
    const aanbodText = row[4];
    if (aanbodText && typeof aanbodText === 'string' && currentNr) {
      const isJa = row[5] === true || String(row[5]).toLowerCase() === 'true';
      // Store by exact Aanbod string so we can match it in the dashboard
      data.aanbod[aanbodText.trim()] = isJa;
      if(isJa) {
        n++;
        count++;
      }
    }
  }
  if(count === 0) item.errors.push('Aanbod niet ingevuld');
  data.status[2] = n;

  // --- TAB 4: Samenwerking ---
  const tab4 = getSheetData('4. (provinciale) Samenwerking');
  n = 0;
  count = 0;
  for (let r = 0; r < tab4.length; r++) {
    const row = tab4[r];
    const nr = parseInt(row[1]);
    if (!isNaN(nr)) {
      const zelf = row[4] === true || String(row[4]).toLowerCase() === 'true';
      const samen = row[5] === true || String(row[5]).toLowerCase() === 'true';
      
      let val = 0;
      if (zelf && samen) val = 3;
      else if (zelf) val = 1;
      else if (samen) val = 2;

      data.samenwerking[nr] = {
        waarde: val,
        toelichting: row[6] || ''
      };
      if(val >= 2) n++;
      if(val > 0) count++;
    }
  }
  if(count === 0) item.errors.push('Provinciale samenwerking niet ingevuld');
  data.status[3] = n;

  // --- TAB 5: Maatwerkplan collectief ---
  const tab5 = getSheetData('5. Maatwerkplan');
  n = 0; count = 0;
  if (tab5.length > 0) {
    let headerRowIndex = -1;
    const colIndices = { nr: -1, aanpak: -1, voortgang: -1, status: -1, planning: -1, trekker: -1 };

    // 1. Find the header row and dynamically locate the target columns
    for (let r = 0; r < tab5.length; r++) {
      const row = tab5[r];
      let foundNr = -1;
      
      // Look for the "Nr" column to identify the header row
      for (let c = 0; c < row.length; c++) {
        if (String(row[c]).trim() === 'Nr') foundNr = c;
      }

      if (foundNr !== -1) {
        headerRowIndex = r;
        colIndices.nr = foundNr;
        
        // Now scan the rest of this row for the keyword columns
        for (let c = 0; c < row.length; c++) {
          const cellText = String(row[c] || '').toLowerCase();
          if (cellText.includes('aanpak')) colIndices.aanpak = c;
          if (cellText.includes('voortgang')) colIndices.voortgang = c;
          if (cellText.includes('status')) colIndices.status = c;
          if (cellText.includes('planning')) colIndices.planning = c;
          if (cellText.includes('trekker')) colIndices.trekker = c;
        }
        break; // Stop searching once we found and mapped the headers
      }
    }

    // 2. Loop through the rows below the header and group them by "Nr"
    if (headerRowIndex !== -1) {
      let currentNr = null;
      
      for (let r = headerRowIndex + 1; r < tab5.length; r++) {
        const row = tab5[r];
        const cellNrValue = row[colIndices.nr];
        
        // If there is a number in the Nr column, we enter a new Criterium block
        if (cellNrValue !== undefined && cellNrValue !== null && cellNrValue !== '') {
          const parsedNr = parseInt(cellNrValue);
          if (!isNaN(parsedNr)) {
            currentNr = parsedNr;
            if (!data.maatwerkplan[currentNr]) {
              // Initialize arrays to collect strings
              data.maatwerkplan[currentNr] = { aanpak: [], voortgang: [], status: [], planning: [], trekker: [] };
            }
          }
        }

        // If we are currently inside a valid Criterium block, extract the data
        if (currentNr && data.maatwerkplan[currentNr]) {
          const extractAndDeduplicate = (key) => {
            if (colIndices[key] !== -1) {
              const val = String(row[colIndices[key]] || '').trim();
              
              // Only add if it's not empty AND not already in the array (removes duplicates from merged cells)
              if (val && !data.maatwerkplan[currentNr][key].includes(val)) {
                data.maatwerkplan[currentNr][key].push(val);
              }
            }
          };

          extractAndDeduplicate('aanpak');
          extractAndDeduplicate('voortgang');
          extractAndDeduplicate('status');
          extractAndDeduplicate('planning');
          extractAndDeduplicate('trekker');
        }
      }

      // 3. Finalize by joining the collected arrays into single strings with newlines
      for (const nr in data.maatwerkplan) {
        data.maatwerkplan[nr].aanpak = data.maatwerkplan[nr].aanpak.join('\n');
        data.maatwerkplan[nr].voortgang = data.maatwerkplan[nr].voortgang.join('\n');
        data.maatwerkplan[nr].status = data.maatwerkplan[nr].status.join('\n');
        data.maatwerkplan[nr].planning = data.maatwerkplan[nr].planning.join('\n');
        data.maatwerkplan[nr].trekker = data.maatwerkplan[nr].trekker.join('\n');
      }
    }
  }
  count = Object.keys(data.maatwerkplan).length;
  n = Object.values(data.maatwerkplan).filter(item => !!item.aanpak.trim()).length;
  if(count === 0 || n === 0) item.errors.push(`Maatwerkplan niet ingevuld`);
  data.status[4] = n;

  if(missingSheets.length > 0) {
    data.status[5] = '❌';
    item.errors = missingSheets;
    data.valid = false;
  } else {
    data.status[5] = item.errors.length > 0 ? `⚠️ ${item.errors.length}x` : '✅';
  }

  item.data = data;
  return item;
}