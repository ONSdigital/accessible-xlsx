import { resolve } from "node:path";
import { defineConfig } from "vite";

export default defineConfig({
	build: {
		lib: {
			entry: resolve(import.meta.dirname, "./lib/main.ts"),
			name: "accessibleXLSX",
			fileName: (format) => `accessible-xlsx.${format}.js`
		},
		rolldownOptions: {
			output: { inlineDynamicImports: true }
		}
	}
});
