import { defineConfig } from 'vite';

export default defineConfig(({ command }) => ({
  // GitHub Pages project URL: https://kamlesh-thakur.github.io/KPI-Dash/
  base: command === 'build' ? '/KPI-Dash/' : '/',
  server: {
    port: 3000,
    open: true
  }
}));
