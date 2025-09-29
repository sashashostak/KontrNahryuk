import { defineConfig } from 'vite'

export default defineConfig({
  root: './src',
  base: './',
  build: { outDir: '../dist/renderer', emptyOutDir: true },
  server: { port: 5177 },
})
