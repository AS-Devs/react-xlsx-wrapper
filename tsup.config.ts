import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm', 'cjs'],
  dts: false,          // manual types/index.d.ts is the published API contract for now
  minify: true,
  sourcemap: true,
  clean: true,
  outDir: 'dist',
  external: ['react', 'react-dom', 'xlsx-js-style'],
  esbuildOptions(options) {
    options.jsx = 'automatic';
  },
});
