import { resolve } from "node:path";
import { defineConfig, loadEnv } from "vite";

// Bundle everything into a single file. Tree-shaking is disabled because rolldown otherwise
// splits documonster's excel/surface/worksheet.js into its own chunk (it is imported both
// statically and via documonster's dynamic xlsb imports), which breaks the UMD build.
const rolldownOptions = {
	output: { codeSplitting: false },
	treeshake: false
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
