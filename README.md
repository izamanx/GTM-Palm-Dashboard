# GTM - Palm Dibbling Dashboard

View-only dashboard. It reads `data/palm-data.xlsx` in the browser every time the page loads.

## Updating the data
1. Rename the new Excel file to `palm-data.xlsx`.
2. Replace the file in the `data/` folder on GitHub and commit.
3. Open `script.js` on GitHub and change the `updatedOn` date near the top (shown in the page header).
4. Vercel redeploys automatically.

First sheet needs these column headings (any order): `Year`, `Date`, `District`,
`Department Type`, `Gov Department Type`, `Seedlings`.

## Settings (top of script.js)
- `title` - page heading
- Logos are in `assets/` - replace the two PNG files to change them.

## Vercel
Import the repository, Framework Preset "Other", no build command.

## Testing on your computer
Do not double-click index.html. Run `npx serve` (or `python -m http.server`) in this folder.
