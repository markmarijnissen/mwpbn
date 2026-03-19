import { ref, toRaw, shallowRef, triggerRef, watch } from "vue";
import { set, get } from 'idb-keyval';
import ExcelJS from "exceljs";
import _ from "lodash";

export const files = shallowRef([]);
export const dashboard = ref({
  type: "default",
  filename: "",
  defaultContent: null,
  customContent: null,
  workbook: null
});

export const addFiles = async (file) => {
  const newFiles = file.name.endsWith('.zip') ? await unzip(file) : [ { filename: file.name, content: await file.arrayBuffer()}];
  await Promise.all(newFiles.map(parsePlan));
  files.value = _.sortBy(files.value
    .filter(item => !newFiles.map(i => i.filename).includes(item.filename)) // prevent double entries
    .concat(newFiles), ["data.contactgegevens.collectief","data.modified"]);
  triggerRef(files);
}

export const removeFile = async (index) => {
  files.value.splice(index, 1);
  triggerRef(files);
}

export const removeInvalidFiles = () => {
  files.value = files.value.filter(file => file.data.valid);
  triggerRef(files);
};

export const removeOldVersions = () => {
  files.value = _(files.value)
    .groupBy("data.contactgegevens.collectief")
    .mapValues(values => values[0])
    .values()
    .value();
  triggerRef(files);
};

export const setCustomDashboard = async (file) => {
  dashboard.value.type = "custom";
  dashboard.value.filename = file.name;
  dashboard.value.customContent = await file.arrayBuffer();
}

// Helper
const createDashboardWorkbook = async () => {
  await loaded;
  console.log(`createDashboardWorkbook`);
  const dashboardWb = new ExcelJS.Workbook();
  await dashboardWb.xlsx.load(dashboard.value.type === "custom" ? dashboard.value.customContent : dashboard.value.defaultContent);
  
  // 2.5 Workaround for ExcelJS Conditional Formatting bug
  dashboardWb.worksheets.forEach(sheet => {
    if (sheet.conditionalFormattings) {
      sheet.conditionalFormattings.forEach(cf => {
        if (cf.rules) {
          cf.rules.forEach(rule => {
            // If the rule is missing the formulae array, initialize it
            if (!rule.formulae) {
              rule.formulae = []; 
            }
          });
        }
      });
    }
  });

  // 2. Loop through all uploaded files
  await Promise.all(files.value.filter(file => file.data.valid).map(async (file, i) => {
    await addPlanToDashboard(file, dashboardWb, i + 1);
  }));

  dashboard.value.workbook = markRaw(dashboardWb);
}

// Init
export const loaded = 
  get('files')
    .then(value => Promise.all((value || []).map(parsePlan)))
    .then(value => files.value = value)
    .then(() => get('dashboard'))
    .then(value => {
      Object.assign(dashboard.value,value || {});
      return fetch('/dashboard.xlsx');
    })
    .then(res => res.arrayBuffer())
    .then(value => dashboard.value.defaultContent = value)
    .then(() => console.log('loaded', files.value.length))
    .catch(err => console.error('Loading error', err))

// Reactivity; save in browser & update workbook result
watch(files, value => {
  set('files', value.map(({ filename, content }) => ({ filename, content })));
  createDashboardWorkbook();
});

watch(
  () => [dashboard.value.type, dashboard.value.filename, dashboard.value.customContent], 
  () => {
    const data = toRaw(dashboard.value);
    set('dashboard', { type: data.type, filename: data.filename, customContent: data.customContent });
    createDashboardWorkbook();
  }, 
  { deep: true }
);

// debug
window.get = get; window.set = set; window.dashboard = dashboard; window.files = files;
window.update = async () => {
  await Promise.all(files.value.map(parsePlan));
  createDashboardWorkbook();
}
