import { defineConfig } from 'vite';
import { resolve } from 'path';
import officeCerts from 'office-addin-dev-certs';

// Plugins
import Vue from '@vitejs/plugin-vue';
import AutoImport from 'unplugin-auto-import/vite'


export default defineConfig(async ({ mode }) => {
  // Get HTTPS certificate options
  const httpsOptions = process.env.USE_DEV_CERTS !== 'false' 
    ? await officeCerts.getHttpsServerOptions() 
    : {};

  return {
    plugins: [
      Vue(),
      AutoImport({
        dts: true,
        defaultExportByFilename: true,
        include: [
          /\.[tj]sx?$/, // .ts, .tsx, .js, .jsx
          /\.vue$/,
          /\.vue\?vue/, // .vue
          /\.vue\.[tj]sx?\?vue/, // .vue (vue-loader with experimentalInlineMatchResource enabled)
        ],
        imports: [
          'vue',
        ],
      })
    ],
    base: '',
    build: {
      outDir: 'dist',
    },
    resolve: {
      alias: {
        '@': resolve(__dirname, 'src'),
      },
    },
    server: {
      port: 3000,
      https: httpsOptions,
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
    },
  };
});