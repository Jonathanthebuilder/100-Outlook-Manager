import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  base: './',
  build: {
    rollupOptions: {
      input: {
        'outlook-manager': resolve(__dirname, 'outlook-manager.html'),
      },
    },
  },
  plugins: [react()],
  test: {
    environment: 'node',
    globals: true
  }
});
