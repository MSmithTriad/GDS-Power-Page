# GDS-Power-Page

A Power Page themed with GDS components.

Add `.ts` files with the same name as the generated `.js` files and then run `npm run build:watch` to develop the scripts with TypeScript.

## CLI Commands

- `pac pages list` - To get the list of pages IDs.
- `pac pages download --path [path] --webSiteId [id] -mv Enhanced` - To download the latest page. Replace `path` and `id`.
- `pac pages upload --path [path]  --modelVersion 2` - To upload the page changes.
- `npm run build` - Compile all .ts files to .js.
- `npm run build:watch` - Watch for changes and auto-compile.
