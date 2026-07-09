<template>
  <NuxtRouteAnnouncer />
  <NuxtLayout >
    <Menu class="hide-print"></Menu>
    <div class="content container p-3" data-theme="light">
      <NuxtPage />
    </div>
  </NuxtLayout>
</template>
<style>
@media print {
  .hide-print {
    display: none;
  }
}
</style>
