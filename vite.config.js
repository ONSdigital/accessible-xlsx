import { resolve } from "node:path";
import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";

const name = "accessible-xlsx";

module.exports = defineConfig({
	build: {
		lib: {
			entry: resolve(__dirname, "lib/main.js"),
			name,
			fileName: (format) => `${name}.${format}.js`
		}
	},
	plugins: [nodePolyfills()]
});
