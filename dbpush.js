import { config } from "dotenv";
import { execSync } from "child_process";

config();

execSync("npx drizzle-kit push", { stdio: "inherit", env: process.env });
