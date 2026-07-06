import { resolve } from "node:path";
import { defineConfig, loadEnv } from "vite";

export default defineConfig((mode) => {
	return mode.mode === "docs"
		? {
				build: {
					outDir: "./docs",
					rolldownOptions: {
						output: { inlineDynamicImports: true }
					}
				}
			}
		: {
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
			};
});
