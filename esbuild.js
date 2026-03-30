const esbuild = require('esbuild');

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

async function main() {
  const ctx = await esbuild.context({
    entryPoints: ['src/extension.ts'],
    bundle: true,
    format: 'cjs',
    minify: production,
    sourcemap: !production,
    sourcesContent: false,
    platform: 'node',
    outfile: 'dist/extension.js',
    external: ['vscode'],
    // @qlik/api is ESM-only; esbuild handles ESM→CJS conversion automatically.
    // We keep it bundled so the extension ships as a single file.
    logLevel: 'silent',
    // Suppress "default export in CommonJS" warnings from ESM packages
    banner: {
      js: '"use strict";',
    },
    plugins: [
      {
        name: 'esbuild-problem-matcher',
        setup(build) {
          build.onStart(() => console.log('[watch] build started'));
          build.onEnd(result => {
            result.errors.forEach(({ text, location }) => {
              console.error(`✘ [ERROR] ${text}`);
              if (location) console.error(`    ${location.file}:${location.line}:${location.column}:`);
            });
            if (result.errors.length === 0) console.log('[watch] build finished');
          });
        },
      },
    ],
  });

  if (watch) {
    await ctx.watch();
  } else {
    await ctx.rebuild();
    await ctx.dispose();
  }
}

main().catch(e => {
  console.error(e);
  process.exit(1);
});
