# Word to HTML Studio

A local-first React web app for converting Word `.docx` documents into editable HTML. It uses Mammoth.js to extract document content, embeds DOCX images as base64 data URLs, provides a WordPress Classic Editor-inspired visual editor, and includes real-time HTML cleanup options.

## Features

- Upload a `.docx` file or drag one into the import area.
- Paste rich content copied from Microsoft Word, or paste plain text.
- Embedded DOCX images are converted to `data:image/...;base64,...` URLs.
- Switch between **Live view** and editable **HTML code**.
- Classic-editor-style controls for headings, bold, italic, underline, lists, links, alignment, quotes, and horizontal rules.
- Live cleanup toggles for inline styles, classes/IDs, empty tags, bold, italic, underline, hyperlinks, font/span tags, comments, and broad formatting.
- HTML source formatting adds a blank source line before each heading from `<h1>` through `<h6>`.
- Export a full HTML document, HTML fragment, plain text (`.txt`), Markdown (`.md`), or a JSON bundle (`.json`).
- Copy generated HTML to the clipboard.
- Responsive desktop and mobile layouts.

## Requirements

- Node.js 20.19+ or 22.12+ recommended.
- npm.

## Run locally

```bash
npm install
npm run dev
```

Vite prints a local URL, usually `http://localhost:5173`. Open it in a modern browser.

## Production build

```bash
npm run build
npm run preview
```

The production files are created in `dist/` and can be hosted as a static site.

## Notes and limitations

- Mammoth runs in the browser and supports `.docx`; the legacy binary `.doc` format must first be saved as `.docx` from Microsoft Word or LibreOffice.
- DOCX images are read from the document package and embedded as base64 data URLs. Pasted images are embedded when the browser exposes image data, and reachable external images are fetched and converted where browser CORS permissions allow it.
- Complex Word pagination, floating objects, SmartArt, macros, and precise page layout are not represented exactly in HTML. Mammoth focuses on semantic document content rather than pixel-perfect Word rendering.
- Cleanup options are output transformations. After editing the transformed content in the visual or code editor, those edits become the new source and may no longer be reversible by disabling a cleanup toggle.
- The app performs conversion on the client; document contents are not sent to an application server. Google Fonts are loaded from Google when internet access is available; the app still works without them using system fallbacks.

## Main dependencies

- React + Vite
- Mammoth.js
- DOMPurify
- Turndown
- Tailwind CSS
- Lucide React icons
