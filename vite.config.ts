import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig({
  base: '/espaco-angelica-preview/',
  plugins: [react(), tailwindcss()]
});
