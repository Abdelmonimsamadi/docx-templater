import { build } from "esbuild";
import { execSync } from "child_process";
import { rmSync } from "fs";

// Clean old builds
rmSync("dist", { recursive: true, force: true });

// 1️⃣ Build JavaScript with ESBuild
await build({
    entryPoints: ["src/index.ts"],
    bundle: true,
    format: "esm",         // use "cjs" if you prefer CommonJS
    platform: "neutral",   // "browser" or "node" depending on your lib
    target: "es2020",
    outdir: "dist",
    sourcemap: true,
    minify: false
});

// 2️⃣ Generate Type Definitions
execSync("tsc --emitDeclarationOnly", { stdio: "inherit" });

console.log("✅ Build complete!");
