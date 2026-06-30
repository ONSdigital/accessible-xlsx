import { resolve } from "node:path";
import { defineConfig } from "vite";

module.exports = defineConfig({
	build: {
		lib: {
			entry: resolve(__dirname, "lib/main.js"),
			name: "accessibleXLSX",
			fileName: (format) => `accessible-xlsx.${format}.js`
		}
	}
});
