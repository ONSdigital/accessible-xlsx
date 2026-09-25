import { resolve } from "node:path";
import { defineConfig, loadEnv } from "vite";

// Bundle everything into a single file.
const rolldownOptions = {
	output: { codeSplitting: false }
};

export default defineConfig((mode) => {
	return mode.mode === "docs"
		? {
				build: {
					outDir: "./docs",
					rolldownOptions
				},
				base: ""
			}
		: {
				build: {
					lib: {
						entry: resolve(import.meta.dirname, "./lib/main.ts"),
						name: "accessibleXLSX",
						fileName: (format) => `accessible-xlsx.${format}.js`
					},
					rolldownOptions
				}
			};
});
