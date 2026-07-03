import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    strictPort: true,
    open: true,
    host: true, // listen on all interfaces so tunnels / other devices can reach it
    // Vite blocks unknown Host headers by default. Allow ngrok tunnel domains
    // (a leading dot matches any subdomain, so it survives ngrok's changing URLs).
    allowedHosts: ['.ngrok-free.app', '.ngrok.app', '.ngrok.io'],
    // Same-origin API proxy. When the portal is opened from a phone / ngrok, it
    // calls the API on its own origin (see src/api.js) and this forwards those
    // /api requests to the local service — avoiding the unreachable "localhost"
    // and https→http mixed-content that cause "Failed to fetch" on mobile.
    proxy: {
      '/api': { target: 'http://localhost:4000', changeOrigin: true },
    },
  },
});
