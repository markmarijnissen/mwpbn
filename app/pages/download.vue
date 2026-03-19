<template>
    <h4>Stap 3. Download het dashboard</h4>
    <template v-if="files.length > 0">
        <a href="#" class="button is-primary" @click.prevent="downloadDashboard">Download dashboard</a>
    </template>
    <template v-else>
        <article class="message is-danger">
            <div class="message-body">
                Je hebt nog geen maatwerkplannen geselecteerd.
            </div>
        </article>
        <nuxt-link to="/files" class="button">Selecteer maatwerkplannen</nuxt-link>
    </template>

    <template v-if="dashboard.workbook">
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

await loaded;


</script>