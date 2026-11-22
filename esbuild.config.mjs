import esbuild from "esbuild";
import process from "process";
import { readFileSync } from "node:fs";

const prod = (process.argv[2] === 'production');
const markedLicense = readFileSync(new URL("./licenses/marked.txt", import.meta.url), "utf8");

esbuild.build({
    banner: {
        js: `/* SmartMD 1.0.22 */\n/*! Marked license\n${markedLicense}\n*/`,
    },
    entryPoints: ['main.ts'],
    bundle: true,
    external: [
        'obsidian',
        'electron',
        '@codemirror/autocomplete',
        '@codemirror/collab',
        '@codemirror/commands',
        '@codemirror/language',
        '@codemirror/lint',
        '@codemirror/search',
        '@codemirror/state',
        '@codemirror/view',
        '@lezer/common',
        '@lezer/highlight',
        '@lezer/lr',
        'node:crypto',
        'fs',
        'path',
        'child_process'
    ],
    format: 'cjs',
    target: 'es2018',
    logLevel: "info",
    sourcemap: prod ? false : 'inline',
    treeShaking: true,
    outfile: 'main.js',
}).catch(() => process.exit(1));
