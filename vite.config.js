import { defineConfig } from 'vite';

// https://vitejs.dev/config/
export default defineConfig({
    base: '/larkbase-attachment-downloader/',
    server: {
        host: '0.0.0.0',
    },
    build: {
        target: ['edge90', 'chrome90', 'firefox90', 'safari15']
    }
})
