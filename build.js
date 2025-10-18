import { build } from "esbuild";
import { execSync } from "child_process";
import { rmSync } from "fs";

// Clean old builds
rmSync("dist", { recursive: true, force: true });

// 1️⃣ Build JavaScript with ESBuild
await build({
    entryPoints: ["src/index.ts"],
    bundle: true,
    format: "esm",
    platform: "node",
    outdir: "dist/esm"
});

await build({
    entryPoints: ["src/index.ts"],
    bundle: true,
    format: "cjs",
    platform: "node",
    outdir: "dist/cjs",
    outExtension: { ".js": ".cjs" },
});

// 2️⃣ Generate Type Definitions
execSync("tsc --emitDeclarationOnly", { stdio: "inherit" });

console.log("✅ Build complete!");
