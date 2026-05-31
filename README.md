# gscriptSETHclasp

Existing Google Apps Script project connected with `clasp`.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```

2. Authenticate `clasp` if needed:
   ```bash
   npx clasp login
   ```

3. Pull the connected project:
   ```bash
   npx clasp pull
   ```

4. Push changes back to Apps Script:
   ```bash
   npx clasp push
   ```

## Notes

- The current project is configured in `.clasp.json`.
- Source files live in `src/`.
- `tsconfig.json` and TypeScript support were removed because this project is using JavaScript.
