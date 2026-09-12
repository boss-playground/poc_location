import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";
export default defineConfig({
  base: "/calendar/",
  build: {
    rolldownOptions: {
      input: {
        main: fileURLToPath(new URL("./index.html", import.meta.url)),
        redirect: fileURLToPath(new URL("./redirect.html", import.meta.url)),
        logout: fileURLToPath(new URL("./logout.html", import.meta.url)),
      },
    },
  },
});
