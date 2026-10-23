import { defineConfig } from "vite";

export default defineConfig({
  // 하위 경로에 올려도 동작하도록 상대 경로로 빌드
  base: "./",
  build: { target: "es2022" },
});
