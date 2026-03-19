<template>
  <div class="excel-preview-wrapper is-fullwidth" v-if="sheets.length > 0">
    <p class="mb-1">Preview (zonder opmaak en formules)</p>
    <div class="tabs is-boxed is-small">
      <ul>
        <li 
          v-for="(sheet, index) in sheets" 
          :key="index"
          :class="{ 'is-active': activeTabIndex === index }"
        >
          <a @click="activeTabIndex = index">
            {{ sheet.name }}
          </a>
        </li>
      </ul>
    </div>

    <div class="table-container">
      <table class="table is-bordered is-striped is-narrow is-hoverable is-fullwidth preview-table">
        <tbody>
          <tr v-for="(row, rowIndex) in activeSheet.rows" :key="rowIndex">
            <th class="row-number has-text-grey-light has-background-white-ter">{{ rowIndex + 1 }}</th>
            
            <td v-for="(cell, cellIndex) in row" :key="cellIndex" :title="cell">
              {{ cell }}
            </td>
          </tr>
          
          <tr v-if="activeSheet.rows.length === 0">
            <td class="has-text-centered has-text-grey is-italic py-4">
              Dit tabblad is leeg.
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
  
  <div v-else class="notification is-light has-text-centered">
    Geen voorbeeld beschikbaar of bestand is aan het laden...
  </div>
</template>

<script setup>
import { ref, computed, watch, toRaw } from 'vue';

const props = defineProps({
  // Expects an ExcelJS Workbook instance
  workbook: {
    type: Object,
    required: true
  }
});

const sheets = ref([]);
const activeTabIndex = ref(2);

// Computed property to quickly grab the currently selected sheet's data
const activeSheet = computed(() => {
  return sheets.value[activeTabIndex.value] || { rows: [] };
});

// Watch the workbook prop so if a new file is loaded, the preview updates
watch(() => props.workbook, (newWb) => {
  if (!newWb || !newWb.worksheets) {
    sheets.value = [];
    return;
  }

  const extractedSheets = [];

  // Loop through all worksheets in the ExcelJS workbook
  newWb.worksheets.forEach((sheet) => {
    const rows = [];
    
    // .eachRow guarantees we iterate through rows that have data
    sheet.eachRow((row) => {
      const rowData = [];
      
      // .eachCell with { includeEmpty: true } ensures our columns align perfectly
      // even if there are blank cells in the middle of a row
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // cell.text automatically resolves 
        // formulas to their results and formats dates/numbers cleanly as strings
        rowData[colNumber - 1] = cell.text || ''; 
      });
      
      rows.push(rowData);
    });

    extractedSheets.push({
      name: sheet.name,
      rows: rows
    });
  });

  sheets.value = extractedSheets;
  activeTabIndex.value = 2; // Reset to the first tab when a new file is loaded
}, { immediate: true, deep: true }); // immediate: true runs this as soon as the component mounts
</script>

<style scoped>
.excel-preview-wrapper {
  background-color: #fff;
}

/* Tweak Bulma's default tab margins so it sits flush with the wrapper */
.tabs {
  margin-bottom: 0 !important;
  background-color: #fcfcfc;
  border-bottom: 1px solid #dbdbdb;
  border-radius: 6px 6px 0 0;
}
.tabs.is-boxed a {
  border-radius: 6px 6px 0 0;
  border-bottom: none;
}

.table-container {
  max-height: 500px; /* Gives it a nice scrollable area */
  overflow: auto;
  margin-bottom: 0;
}

/* Make the table look a bit more like a spreadsheet */
.preview-table {
  font-size: 0.85rem;
  white-space: nowrap; /* Prevents cells from wrapping text, creating weird heights */
}

.preview-table td, .preview-table th {
  border: 1px solid #e2e2e2;
  vertical-align: middle;
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
}

.row-number {
  width: 40px;
  text-align: center !important;
  user-select: none;
}
</style>