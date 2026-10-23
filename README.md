# Innomick Technologies - Document Management System (DMS) Prototype

This is the standalone **Document Management System (DMS)** prototype for Innomick Technologies.

## 🚀 How to Run

You can run this prototype independently using any of the following methods:

### Method 1: Using Node.js (Recommended)
```bash
# Navigate to the folder if you are not already in it
npm start
# or
node server.js
```
The prototype will be live at:
👉 **[http://localhost:3000](http://localhost:3000)**

*(To use a custom port, set the PORT environment variable: `PORT=8080 node server.js`)*

### Method 2: Direct Browser File (No server required)
Simply double-click [`index.html`](index.html) or open it directly in Google Chrome, Microsoft Edge, or any modern web browser.

### Method 3: Live Server / Any HTTP Server
You can serve this folder with any static web server:
```bash
npx serve .
# or VS Code Live Server extension
```

---

## 📁 File Structure
- [`index.html`](index.html) - Main SharePoint-styled DMS application interface
- [`style.css`](style.css) - Enterprise Microsoft 365 / SharePoint styling & layouts
- [`app.js`](app.js) - Interactive behavior, filters, multi-step upload workflow, search, and modals
- [`data.js`](data.js) - Preloaded document libraries, organizational hierarchy, users, workflows, and mock datasets
- `assets/` - Brand logos (`innomick_logo_horizontal.png`, `innomick_logo_square.png`)
- `screenshots/` - Captured reference screenshots and feature walkthroughs
- `server.js` - Zero-dependency Node.js HTTP web server
- `package.json` - Node project scripts (`npm start`)
