import { dashboard } from "~/services/files";

export async function downloadDashboard() {
  dashboard.value.workbook.calcProperties.fullCalcOnLoad = true;
  // Force OnlyOffice/LibreOffice to recalculate by destroying all cached formula results
  dashboard.value.workbook.worksheets.forEach(sheet => {
    sheet.eachRow(row => {
      row.eachCell(cell => {
        // If the cell contains a formula, we rewrite it WITHOUT the 'result' property
        if (cell.formula || cell.sharedFormula) {
          cell.value = {
            formula: cell.formula,
            sharedFormula: cell.sharedFormula
            // Notice we intentionally leave 'result' undefined!
          };
        }
      });
    });
  });

  const updatedDashboardBuffer = await dashboard.value.workbook.xlsx.writeBuffer();
  const blob = new Blob([updatedDashboardBuffer], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', 'dashboard.xlsx'); 
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}