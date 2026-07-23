import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { VitePWA } from 'vite-plugin-pwa';

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      // Precache the whole app shell: the app must launch with zero connectivity.
      workbox: {
        globPatterns: ['**/*.{js,css,html,ico,png,svg,woff2}'],
        navigateFallback: '/index.html',
        // Never let the SW intercept flow calls — sync logic owns its own retries.
        navigateFallbackDenylist: [/^\/api/],
      },
      manifest: {
        name: 'Asset Stocktake',
        short_name: 'Stocktake',
        description: 'Offline-first asset stocktake with Excel Online reconciliation',
        start_url: '/',
        display: 'standalone',
        orientation: 'portrait',
        background_color: '#101418',
        theme_color: '#101418',
        icons: [
          { src: '/icon-192.png', sizes: '192x192', type: 'image/png' },
          { src: '/icon-512.png', sizes: '512x512', type: 'image/png' },
          { src: '/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
        ],
      },
    }),
  ],
});
