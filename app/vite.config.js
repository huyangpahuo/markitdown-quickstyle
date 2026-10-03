import { defineConfig } from 'vite';

export default defineConfig({
  clearScreen: false,
  server: {
    port: 5173,
    strictPort: true,
  },
  build: {
    target: 'chrome105',
    rollupOptions: {
      input: {
        main: 'index.html',
        pet: 'pet.html',
      },
    },
  },
});
