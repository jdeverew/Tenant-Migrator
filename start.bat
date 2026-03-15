@echo off
echo Installing / updating dependencies...
npm install
echo.
echo Starting application...
set NODE_ENV=development
npx tsx server/index.ts
