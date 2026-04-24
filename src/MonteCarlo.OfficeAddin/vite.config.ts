import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    host: 'localhost',
    port: 3000,
    https: true,
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'taskpane.html'),
        'custom-functions': resolve(__dirname, 'src/custom-functions.ts'),
      },
      output: {
        entryFileNames: 'assets/[name].js',
        chunkFileNames: 'assets/[name].js',
        assetFileNames: 'assets/[name][extname]',
      },
    },
  },
});
