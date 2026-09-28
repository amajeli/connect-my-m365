import { build as esbuild } from "esbuild";
import { build as viteBuild } from "vite";
import { rm } from "fs/promises";

// Netlify build: the client goes to dist/public (served by the CDN) and the
// Express API is bundled into one self-contained function file.
async function buildNetlify() {
  await rm("dist", { recursive: true, force: true });
  await rm("netlify/functions-dist", { recursive: true, force: true });

  console.log("building client...");
  await viteBuild();

  console.log("building function...");
  await esbuild({
    entryPoints: ["netlify/functions/api.ts"],
    platform: "node",
    target: "node20",
    bundle: true,
    // ESM because package.json has "type": "module"; bundled CommonJS
    // dependencies still get a working require() through the banner.
    format: "esm",
    outfile: "netlify/functions-dist/api.mjs",
    banner: { js: "import { createRequire as __cr } from 'module'; const require = __cr(import.meta.url);" },
    define: { "process.env.NODE_ENV": '"production"' },
    // Optional native add-ons that pg and ws fall back from when absent.
    external: ["pg-native", "bufferutil", "utf-8-validate"],
    minify: true,
    logLevel: "info",
  });
}

buildNetlify().catch((err) => {
  console.error(err);
  process.exit(1);
});
