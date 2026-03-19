
<template>
    <h4>Stap 2. Gebruik eigen dashboard (optioneel)</h4>
        <article class="message">
        <div class="message-body">
        <p>De maatwerkplannen worden samengevoegd in één Excel bestand: <a href="/dashboard.xlsx">het dashboard</a>.</p>

        <p>Je kunt dit dashboard <a href="/dashboard.xlsx">downloaden</a> en naar wens aanpassen.</p>
        </div>
        </article>
    <div class="select">
        <select v-model="dashboard.type">
            <option value="default">Gebruik standaard dashboard</option>
            <option value="custom">Gebruik eigen dashboard</option>
        </select>
    </div>

    <div v-if="dashboard.type === 'custom'">
        <FileUpload type="dashboard" class="mt-3"></FileUpload>

        <p class="mt-3">{{ dashboard.customContent ? '✅ '+dashboard.filename : '⚠️ Selecteer bestand' }}</p>
    </div>

</template>
<script setup lang="ts">
import { dashboard } from '~/services/files';
import FileUpload from '~/components/Fileupload.vue';

</script>