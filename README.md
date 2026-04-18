# Org Chart Visualizer

Single-page org chart editor built with React, TypeScript, Vite, and React Flow. It supports hierarchy editing, location grouping, drag-based reassignment, collapse/expand, save/load, and a denser light view for scanning large teams.

## Current App Features

- Org view with reporting lines
- Location view grouped by `Location`
- Full card view and light view
- View mode and drag edit mode
- Drag a card onto a valid manager to reassign reporting
- Drag a manager across sibling managers to reorder whole teams left/right
- Collapse/expand for branches and global collapse/expand controls
- Search by name, role, title, manager/IC, level, worker type, or location
- Side panel editing for the selected card
- Save and load JSON snapshots
- Import `.xlsx`, `.xls`, and `.csv` roster files in-browser
- Export the current view as a landscape PNG image
- Publish-safe default demo org included

## Data Model

Each person record uses the following fields:

- `name`
- `role`
- `managerOrIc`
- `workerType`
- `title`
- `managerName`
- `level`
- `location`
- `roleType`

## Local Development

```bash
npm install
npm run dev
```

Open `http://localhost:5173`.

## Production Build

```bash
npm run build
```

To preview the production build locally:

```bash
npm run preview
```

## Importing Spreadsheet Data

The app includes an `Import spreadsheet` button for non-technical users who prefer Excel or CSV over editing JSON manually.

Supported file types:

- `.xlsx`
- `.xls`
- `.csv`

Expected columns:

- `Name`
- `Role`
- `Manager Or IC`
- `Full Time or Contractor`
- `Title`
- `Manager`
- `Level`
- `Location`

Import behavior:

- the first sheet is used for Excel files
- rows are converted into the app's org JSON format in-browser
- manager relationships are built from the `Manager` column
- duplicate names get unique internal IDs automatically
- if a manager name is not found in the file, that person is attached to the detected root
- the app shows a preview summary before you load the imported org into the canvas

You can also download the converted JSON from that preview if you want a reusable saved file.

## Save / Load Format

The app `Load` button expects a JSON file.

Best option:
- Use the app `Save` button, then later load that same file back with `Load`.

Accepted shapes:

1. Saved snapshot format

```json
{
  "version": 1,
  "savedAt": "2026-04-18T16:20:00.000Z",
  "data": {
    "rootId": "avery-collins",
    "people": []
  }
}
```

2. Raw org data format

```json
{
  "rootId": "avery-collins",
  "people": []
}
```

## Image Export

The `Print image` button exports the current canvas view as a PNG sized for PowerPoint-style landscape slides.

Behavior:

- works with `Org view` and `Location view`
- respects the current card density, so `Light view` exports as light view
- fits all currently visible cards into frame before capture
- restores your previous zoom/pan after export

This is useful for dropping the current org snapshot directly into presentations.

## Default Demo Data

The app boots from [src/data/mockOrg.ts](/Users/johnvincent/Org%20Chart%20Visualizer/src/data/mockOrg.ts), which currently points at the publish-safe sample file:

- [src/data/publicDemoOrg.json](/Users/johnvincent/Org%20Chart%20Visualizer/src/data/publicDemoOrg.json)

That file contains fictional names and is safe to keep in a public repository.

If you want to use your private org data:

1. Keep private data out of `src/` when publishing publicly.
2. Start the app with the public demo org.
3. Use the `Load` button to import your private JSON locally when needed.

## Online Publishing

This app is a static frontend. Any host that can serve the `dist/` folder will work.

General flow:

```bash
npm install
npm run build
```

Deploy the generated `dist/` folder to your host.

Good hosting options:

- GitHub Pages
- Netlify
- Vercel
- Cloudflare Pages

### Cloudflare Pages

This repo is set up to work on Cloudflare Pages while staying on Vite 5.

Use these Cloudflare dashboard settings:

- Framework preset: `Vite`
- Build command: `npm run build`
- Build output directory: `dist`
- Root directory: leave empty unless your repo is nested
- Environment variable changes: none required for the current app

Important:

- Do not set a deploy command like `npx wrangler deploy`
- Do not use the generic Workers deploy flow for this repo

Why:

- this app is a static frontend
- the Vite build already succeeds and produces `dist/`
- the failure you saw comes from Wrangler trying to auto-configure a Worker/Vite deploy path that requires Vite 6+

If Cloudflare currently has a deploy command configured, remove it and let Pages publish the built `dist/` output directly.

If you want to deploy from the CLI instead of Git integration, use the Pages command, not the Workers command:

```bash
npx wrangler pages deploy dist
```

The included [wrangler.toml](/Users/johnvincent/Org%20Chart%20Visualizer/wrangler.toml) is for the Pages workflow and points Cloudflare at `./dist`.

### Publishing Safely

For public hosting:

- Keep only fictional/demo data in the repo default seed.
- Do not commit private roster files into `src/`.
- Test private org data through the `Load` button instead of baking it into the bundle.

### GitHub Pages Note

If you publish to GitHub Pages under a repository path such as:

- `https://username.github.io/repo-name/`

then Vite usually needs a `base` setting in `vite.config.ts`:

```ts
export default defineConfig({
  base: "/repo-name/",
  plugins: [react()]
});
```

If you publish to a root domain such as:

- `https://username.github.io/`

then the default Vite config is usually fine.

## Interaction Notes

- Click a card to edit it in the left panel.
- In `View mode`, the canvas is for browsing.
- In `Drag edit`, left-drag edits cards and middle mouse drag pans the canvas.
- In Org view, ICs and open roles must remain under a valid manager.
- In Location view, dragging a person into another location column updates that person’s `location`.
- `Undo` and `Redo` step through edit history.
- `Light view` compresses cards for high-density scanning.

## Architecture

- [src/App.tsx](/Users/johnvincent/Org%20Chart%20Visualizer/src/App.tsx): app shell, toolbar, editor, drag/drop behavior, save/load flow
- [src/data/mockOrg.ts](/Users/johnvincent/Org%20Chart%20Visualizer/src/data/mockOrg.ts): default org data entry point
- [src/data/publicDemoOrg.json](/Users/johnvincent/Org%20Chart%20Visualizer/src/data/publicDemoOrg.json): public sample load file
- [src/types.ts](/Users/johnvincent/Org%20Chart%20Visualizer/src/types.ts): shared types
- [src/utils/org.ts](/Users/johnvincent/Org%20Chart%20Visualizer/src/utils/org.ts): normalization, layout, filtering, and JSON serialization
- [src/styles.css](/Users/johnvincent/Org%20Chart%20Visualizer/src/styles.css): UI and canvas styling
