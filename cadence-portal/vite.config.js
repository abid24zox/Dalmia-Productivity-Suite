import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // COOP must allow popups or MSAL's sign-in popup floods the console with
  // "Cross-Origin-Opener-Policy would block the window.closed call" errors.
  server: {
    port: 5173,
    strictPort: true,
    open: true,
    // Let an ngrok tunnel (or any host) reach the dev server when sharing.
    allowedHosts: true,
    headers: { 'Cross-Origin-Opener-Policy': 'same-origin-allow-popups', 'Cross-Origin-Embedder-Policy': 'unsafe-none' },
    // One tunnel covers everything: /api/* is proxied to the Cadence service so
    // a remote browser never needs to reach localhost:4000 directly.
    proxy: { '/api': { target: 'http://localhost:4000', changeOrigin: true } },
  },
});
