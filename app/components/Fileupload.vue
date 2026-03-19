<template>
  <div class="file-upload-wrapper">
    <div
      class="box file-drop-zone mb-4"
      :class="{ 'is-dragover': isDragging, 'has-error': errorMessage }"
      @dragenter.prevent="isDragging = true"
      @dragover.prevent="isDragging = true"
      @dragleave.prevent="isDragging = false"
      @drop.prevent="handleDrop"
    >
      <div class="file is-boxed is-centered is-primary has-name">
        <label class="file-label">
          <input
            ref="fileInput"
            class="file-input"
            type="file"
            name="upload"
            :accept="type === 'plan' ? '.xlsx,.zip' : '.xlsx' "
            :multiple="type === 'plan' ?  true : null"
            @change="handleFileChange"
          />
        Kies bestand of sleep bestanden hier naar toe...
        </label>
      </div>
    </div>
    
    <p v-if="errorMessage" class="help is-danger has-text-centered mt-2">
      {{ errorMessage }}
    </p>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { addFiles, setCustomDashboard } from '~/services/files'

const props = defineProps({
  type: {
    type: String,
    default: "plan"
  }
});

const fileInput = ref();
const isDragging = ref(false);
const errorMessage = ref('');
const allowedExtensions = props.type === 'plan' ? ['.xlsx', '.zip'] : ['.xlsx' ];

// Validate file extension manually
const isValidFile = (file) => {
  if (!file) return false
  const fileName = file.name.toLowerCase()
  return allowedExtensions.some(ext => fileName.endsWith(ext))
}

const processFile = (file) => {
  if (isValidFile(file)) {
    if(props.type === 'plan') {
      addFiles(file);
    } else {
      setCustomDashboard(file);
    }
  } else {
    errorMessage.value = 'Ongeldig bestand. Kies ' + allowedExtensions.join(' of ');
  }
}

const handleFileChange = (event) => {
  errorMessage.value = '';
  [ ...event.target.files ].forEach(processFile);
  if (fileInput.value) {
    fileInput.value.value = ''; 
  }
}

const handleDrop = (event) => {
  isDragging.value = false;
  errorMessage.value = '';
  [ ...event.dataTransfer.files ].forEach(processFile);
}
</script>

<style scoped>
.file-drop-zone {
  border: 2px dashed #dbdbdb;
  background-color: #fcfcfc;
  transition: all 0.2s ease-in-out;
  cursor: pointer;
}

/* Hover/Drag state styling */
.file-drop-zone:hover,
.file-drop-zone.is-dragover {
  border-color: var(--bulma-primary); /* Bulma Primary Color */
  background-color: #f5fcf9;
}

/* Error state styling */
.file-drop-zone.has-error {
  border-color: var(--bulma-danger); /* Bulma Danger Color */
  background-color: #feecf0;
}

/* Make sure the label takes up the whole box so clicking anywhere triggers the file dialog */
.file-drop-zone .file-label {
  width: 100%;
  justify-content: center;
}
</style>