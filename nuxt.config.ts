// https://nuxt.com/docs/api/configuration/nuxt-config
export default defineNuxtConfig({
  ssr: false,
  compatibilityDate: '2025-07-15',
  devtools: { enabled: false },

  css: [
    '~/assets/bulma.min.css'
  ],

  app: {
    head: {
      htmlAttrs: {
        'data-theme': 'light'
      }
    }
  },

  modules: ['@nuxt/content']
})