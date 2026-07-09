<template>
    <h4>Doorzoek de plannen</h4>
    <template v-if="files.length > 0">
        <article class="message">
        <div class="message-body hide-print">
            <p>
                Zoek op één of meerdere woorden. 
                Gebruik nummers om te filteren op criteria. 
                Gebruik "ontwikkelopgave" en "voldoet" om daarop te filteren.
            </p>
        </div>
        </article>
        <input type="text" class="input is-rounded" v-model="searchInput" placeholder="Zoekwoorden" />
        <details class="p-2">
            <summary>{{collectieven.length }} collectieven</summary>
            <ol><li v-for="collectief in collectieven" :bind="collectief">{{ collectief}}</li></ol>
        </details>
        <table>
            <thead>
                <tr>
                    <th>Collectief ({{ collectieven.length }}x)</th>
                    <th>Criteria</th>
                    <th>Vraag</th>
                    <th>Antwoord ({{ results.length }}x)</th>
                    <th>Score</th>
                </tr>
            </thead>
            <tr v-for="(result,i) in results" :bind="i">
                <td>{{ result.name }}</td>
                <td :title="criteria[result.nr-1]">{{ result.nr }}</td>
                <td>{{ result.key }}</td>
                <td>{{ result.value }}</td>
                <td>{{ result.score }}</td>
            </tr>
        </table>
    </template>
    <template v-else>
        <article class="message is-danger">
            <div class="message-body">
                Je hebt nog geen maatwerkplannen geselecteerd.
            </div>
        </article>
        <nuxt-link to="/files" class="button">Selecteer maatwerkplannen</nuxt-link>
    </template>
</template>
<style scoped>
.panel-heading {
    margin-bottom: 0;
}
</style>
<script setup lang="ts">
import { files, loaded, database, criteria } from '~/services/files';
import _ from "lodash";
await loaded;

const searchInput = ref(localStorage.search || "");

const results = computed(() => {
    const words = searchInput.value
        .split(" ")
        .map(word => word.trim())
        .filter(word => word.length >= 3 || (isFinite(word) && Number(word) >= 1 && Number(word) <= 42))
        .filter(word => ['voldoet','ontwikkelopgave'].indexOf(word) < 0);
        
    const mustOntwikkelopgave = searchInput.value.indexOf('ontwikkelopgave') >= 0;
    const mustVoldoet = searchInput.value.indexOf('voldoet') >= 0;
    localStorage.search = searchInput.value;

    if(words.length === 0 && !mustVoldoet && !mustOntwikkelopgave) return database.value;

    let results = database.value;
    if(words.length > 0) results = database.value.filter(i => words.some(word => {
        if(isFinite(word)) return i.nr === Number(word);
        return String(i.value).indexOf(word) >= 0 || searchInput.value.indexOf(String(i.name)) >= 0;
    }));
    if(mustVoldoet) results = results.filter(i => i.score === null || i.score.indexOf("voldoet") >= 0);
    if(mustOntwikkelopgave) results = results.filter(i => i.score === null || i.score.indexOf("ontwikkelopgave") >= 0);
    return results;
});

const collectieven = computed(() => _.uniq(results.value.map(i => i.name)));

</script>