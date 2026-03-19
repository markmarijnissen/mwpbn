<template>
    <h4>Stap 3. Download het dashboard</h4>
    <template v-if="files.length > 0">
        <a href="#" class="button is-primary" @click.prevent="downloadDashboard">Download dashboard</a>
        <br/>
        <button v-if="!preview" class="button is-small mt-3" @click.prevent="preview = true">Bekijk preview</button>
    </template>
    <template v-else>
        <article class="message is-danger">
            <div class="message-body">
                Je hebt nog geen maatwerkplannen geselecteerd.
            </div>
        </article>
        <nuxt-link to="/files" class="button">Selecteer maatwerkplannen</nuxt-link>
    </template>

    <template v-if="preview && dashboard.workbook && files.length > 0">
        <ExcelPreview :workbook="dashboard.workbook" class="mt-5"/>
    </template>
</template>
<style scoped>
.panel-heading {
    margin-bottom: 0;
}
</style>
<script setup lang="ts">
import { files, loaded, dashboard } from '~/services/files';
const preview = ref(false);

await loaded;


</script>