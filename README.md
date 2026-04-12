# Layout Toolkit for Google Docs

A free, open-source Google Docs add-on that brings desktop publishing features to your documents. Create professional brochures, newsletters, event programs, business cards, and flyers with ready-made templates, custom branding, and built-in folding guides.

Companion project to [Layout Toolkit for Word](https://github.com/alyssaagard/layout-toolkit).

## Features

**Template Gallery**: Six ready-to-use layouts (tri-fold brochure, bi-fold brochure, classic newsletter, event program, business card, event flyer). Each generates a complete Google Doc with placeholder content formatted with your brand settings.

**Layout Tools**: Direct control over page orientation, paper size, margins, and insertable elements including text placeholders, horizontal dividers, pull quotes, and sidebar callout boxes.

**Brand Profiles**: Save and switch between organization-specific color schemes and font pairings. Brand settings are applied automatically when generating any template. Profiles are stored per user via Google Apps Script PropertiesService.

**Folding and Print Guide**: Built-in visual diagrams showing how to fold tri-fold and bi-fold brochures.

## Installation

### For Users (Direct Setup)

1. Open any Google Doc.
2. Click **Extensions > Apps Script**.
3. Delete the default `myFunction` code in `Code.gs`.
4. Copy the contents of `Code.gs` from this project and paste it in.
5. Click the **+** next to Files and select **HTML**. Name the file `Sidebar` (case sensitive) (without the .html extension, as Apps Script adds it automatically). 
6. Delete the default HTML content and paste in the contents of `Sidebar.html` from this project.
7. Save both files (Ctrl+S / Cmd+S).
8. Close the Apps Script editor tab.
9. Reload the Google Doc.
10. You will see a new **Layout Toolkit** menu in the menu bar.

### For Developers (clasp)

```bash
npm install -g @google/clasp
clasp login
clasp create --type docs --title "Layout Toolkit"
clasp push
```

Then open the linked Google Doc and reload.

## Usage

1. Open the sidebar: **Layout Toolkit > Open Sidebar**
2. (Optional) Go to the **Branding** tab, set your organization name, colors, and fonts, then click **Save Profile** and **Set Active**.
3. Go to the **Templates** tab and click any template card.
4. The document content will be replaced with the selected template, formatted with your active brand.
5. Edit the placeholder text directly in the document.
6. Use the **Layout** tab to adjust page setup, margins, or insert elements like pull quotes and dividers.

## Architecture

```
layout-toolkit-gdocs/
  Code.gs         Server-side Apps Script (menu, templates, branding, layout)
  Sidebar.html    Client-side UI (three-tab sidebar with template gallery)
  README.md       This file
```

## Differences from the Word Version

| Feature | Word Add-in | Google Docs Add-on |
|---|---|---|
| Column layouts | OOXML injection for true columns | Table-based layouts (Docs API limitation) |
| Brand storage | localStorage in browser | PropertiesService (per-user, server-side) |
| Sidebar width | Configurable | Fixed at 300px |
| Page setup | Full Office JS API | DocumentApp (orientation, size, margins) |
| Distribution | Sideload manifest or AppSource | Apps Script editor or Google Workspace Marketplace |

## License

MIT

## Author

Alyssa I. Agard
