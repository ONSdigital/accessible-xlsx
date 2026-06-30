import { resolve } from "node:path";
import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";

module.exports = defineConfig({
	build: {
		lib: {
			entry: resolve(__dirname, "lib/main.js"),
			name: "accessibleXLSX",
			fileName: (format) => `accessible-xlsx.${format}.js`
		}
	},
	plugins: [nodePolyfills()]
});
