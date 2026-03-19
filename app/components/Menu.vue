<template>
  <div class="tabs is-centered mb-0">
    <ul>
      <NuxtLink to="/" custom v-slot="{ navigate, isActive }">
        <li :class="{ 'is-active': isActive }">
          <a @click="navigate" :href="null">Start</a>
        </li>
      </NuxtLink>

      <NuxtLink to="/files" custom v-slot="{ navigate, isActive }">
        <li :class="{ 'is-active': isActive }">
          <a @click="navigate" :href="null">1. Maatwerkplannen</a>
        </li>
      </NuxtLink>

      <NuxtLink to="/custom-dashboard" custom v-slot="{ navigate, isActive }">
        <li :class="{ 'is-active': isActive }">
          <a @click="navigate" :href="null">2. Dashboard</a>
        </li>
      </NuxtLink>

      <NuxtLink to="/download" custom v-slot="{ navigate, isActive }">
        <li :class="{ 'is-active': isActive }">
          <a @click="navigate" :href="null">3. Download</a>
        </li>
      </NuxtLink>
    </ul>
  </div>
</template>