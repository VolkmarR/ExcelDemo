import { defineConfig } from 'vite'
import react, { reactCompilerPreset } from '@vitejs/plugin-react'
import babel from '@rolldown/plugin-babel'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    babel({ presets: [reactCompilerPreset()] })
  ],
  server: {
    proxy: {
      // Forward API calls to the ASP.NET backend (http profile in launchSettings.json)
      '/api': 'http://localhost:5269',
    },
  },
  optimizeDeps: {
    // Do NOT pre-bundle the Univer ESM packages: bundling the graph into one optimized
    // entry produces a ~10 MB file that overflows Vite 8's Rolldown/oxc parser (bogus
    // "invalid JS/JSX" 500 in dev, WASM OOM in build). Excluded, Univer's small compiled
    // ESM files are served directly and parse fine.
    exclude: [
      '@univerjs/presets',
      '@univerjs/preset-sheets-core',
      '@univerjs/preset-sheets-drawing',
      '@univerjs/docs-drawing',
      '@univerjs/drawing',
      '@univerjs/drawing-ui',
      '@univerjs/sheets-drawing',
      '@univerjs/sheets-drawing-ui',
      '@univerjs/core',
      '@univerjs/design',
      '@univerjs/docs',
      '@univerjs/docs-ui',
      '@univerjs/engine-formula',
      '@univerjs/engine-render',
      '@univerjs/network',
      '@univerjs/rpc',
      '@univerjs/sheets',
      '@univerjs/sheets-formula',
      '@univerjs/sheets-formula-ui',
      '@univerjs/sheets-numfmt',
      '@univerjs/sheets-numfmt-ui',
      '@univerjs/sheets-ui',
      '@univerjs/ui',
      '@univerjs/themes',
    ],
    // ...but every leaf dep Univer imports must be optimized so its CommonJS/UMD
    // packages expose proper ESM named/default exports. (.npmrc hoists them so these
    // names resolve from the project root.)
    include: [
      // Jspreadsheet CE: its core dist is a UMD bundle that require()s jsuites and
      // @jspreadsheet/formula, so pre-bundle the whole set for clean ESM interop.
      '@jspreadsheet-ce/react',
      'jspreadsheet-ce',
      '@jspreadsheet/formula',
      'jsuites',
      '@flatten-js/interval-tree',
      '@floating-ui/dom',
      '@floating-ui/utils',
      '@radix-ui/react-dialog',
      '@radix-ui/react-direction',
      '@radix-ui/react-dropdown-menu',
      '@radix-ui/react-hover-card',
      '@radix-ui/react-popover',
      '@radix-ui/react-separator',
      '@radix-ui/react-slot',
      '@wendellhu/redi',
      'async-lock',
      'cjk-regex',
      'class-variance-authority',
      'clsx',
      'decimal.js',
      'fast-diff',
      'franc-min',
      'kdbush',
      'localforage',
      'lodash-es',
      'nanoid',
      'numfmt',
      'opentype.js',
      'ot-json1',
      'prop-types',
      'rbush',
      'react-transition-group',
      'sonner',
      'tailwind-merge',
    ],
  },
})
