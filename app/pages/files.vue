
<template>
    <h4>Stap 1. Selecteer maatwerkplannen</h4>
    <article class="message">
        <div class="message-body">
        <a href="https://boerennatuurnl.sharepoint.com/sites/DoorontwikkelingBoerenNatuur/Gedeelde%20documenten/Forms/AllItems.aspx?id=%2Fsites%2FDoorontwikkelingBoerenNatuur%2FGedeelde%20documenten%2FFase%202%20%2D%20Maatwerkplannen%2F1%2E%20Maatwerkplannen%2FMaatwerkplannen%20collectieven&viewid=60e6f717%2D647a%2D4a22%2Db25d%2Db07ead4d0685&p=true" target="_blank">Download de hele map</a> als zip-bestand en voeg deze toe. Je kunt de maatwerkplannen ook één voor één selecteren.
        </div>
    </article>
    <Fileupload type="plan"></Fileupload>

    <article class="message is-danger" v-if="hasInvalidFiles">
        <div class="message-body">
            <p>Sommige Excelbestanden bevatten geen maatwerkplan.
            <button class="button is-danger is-small ml-3" @click.prevent="removeInvalidFiles">Verwijderen</button>
            </p>
        </div>
    </article>

    <article class="message is-danger" v-if="hasOldVersions && !keep">
        <div class="message-body">
            <p>Sommige collectieven hebben meer dan één maatwerkplan. Wat wil je doen met oude versies?
            <button class="button is-danger is-small ml-3" @click.prevent="removeOldVersions">Verwijderen</button>
            <button class="button is-small ml-3" @click.prevent="keep = true">Bewaren</button>
            </p>
        </div>
    </article>

    <div>
        <table class="is-size-7">
            <thead>
                <tr>
                    <th>#</th>
                    <th>Bestand</th>
                    <th>Gewijzigd</th>
                    <th>Collectief</th>
                    <th title="Contactgegevens (betrokkenen)">1</th>
                    <th title="Ontwikkelpunten">2</th>
                    <th title="Aanbod">3</th>
                    <th title="Provinciale samenwerking">4</th>
                    <th title="Maatwerkplan">5</th>
                    <th>Status</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
                <tr v-for="(file,i) in files" :key="file.name">
                    <td>{{ i + 1 }}</td>
                    <td>{{ file.filename  }}</td>
                    <td>{{ formatDate(file.data?.modified) }}</td>
                    <td>{{ file.data?.contactgegevens.collectief }}</td>
                    <td>{{ file.data?.status[0]}}</td>
                    <td>{{ file.data?.status[1]}}</td>
                    <td>{{ file.data?.status[2]}}</td>
                    <td>{{ file.data?.status[3]}}</td>
                    <td>{{ file.data?.status[4]}}</td>
                    <td :title="file.errors.join('\n')">{{ file.data?.status[5]}}</td>
                    <td><button class="button is-small is-danger" @click="removeFile(i)">Verwijderen</button></td>
                </tr>
            </tbody>
        </table>
        <ul>
        </ul>
    </div>
</template>
<script setup lang="ts">
import Fileupload from '~/components/Fileupload.vue';
import { files, removeInvalidFiles, removeOldVersions } from "../services/files";
import dayjs from 'dayjs';
import _ from "lodash";

function formatDate(d) {
    return dayjs(d).format("DD MMM YYYY (HH:mm)");
}

const keep = ref(false);
watch(files, () => keep.value = false);

const hasOldVersions = computed(() => {
    return _(files.value)
        .groupBy("data.contactgegevens.collectief")
        .mapValues(values => values.length)
        .values()
        .filter(n => n > 1)
        .value()
        .length > 0;
});

const hasInvalidFiles = computed(() => files.value.filter(file => !file.data?.valid).length > 0);
</script>