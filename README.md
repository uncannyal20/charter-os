# Charter OS — Product Charter Intelligence Platform

An AI-powered web application that helps Product Office teams build and maintain Product Charters collaboratively.

## Features
- 👥 Team setup & Terms of Reference generation
- 🎯 Problem Statement (upload files, AI generates draft)
- 🔭 Vision & Strategy (AI generates from inputs)
- 📈 Value Realisation Outcomes (KPIs + Cost Benefit Analysis)
- 🗺 Product Roadmap (phased plan + Gantt chart)
- 📄 Export as PDF, HTML or PowerPoint presentation

---

## Deploy to Vercel (5 minutes)

### 1. Fork or push this repo to GitHub

### 2. Connect to Vercel
1. Go to [vercel.com](https://vercel.com) and sign in with GitHub
2. Click **Add New Project**
3. Select this repository
4. Click **Deploy**

### 3. Add your Anthropic API key
1. In Vercel dashboard → your project → **Settings** → **Environment Variables**
2. Add: `ANTHROPIC_API_KEY` = `sk-ant-your-key-here`
3. Click **Redeploy**

That's it — your team gets a live URL like `https://charter-os.vercel.app`

---

## Run Locally (for development)

### Prerequisites
- Node.js 18+
- Anthropic API key from [console.anthropic.com](https://console.anthropic.com)

### Setup
```bash
# Install dependencies
npm install

# Set API key (Mac/Linux)
export ANTHROPIC_API_KEY=sk-ant-your-key-here

# Set API key (Windows)
set ANTHROPIC_API_KEY=sk-ant-your-key-here

# Start local server
node server.js

# Open in browser
# http://localhost:3131
```

---

## Project Structure

```
charter-os/
├── public/
│   └── index.html          # Main application
├── api/
│   ├── claude.js           # Anthropic API proxy (Vercel serverless)
│   └── generate-slides.js  # PPTX generation (Vercel serverless)
├── server.js               # Local development server
├── vercel.json             # Vercel deployment config
├── package.json
└── README.md
```

---

## Data Storage
All charter data is saved in the browser's localStorage.
Use the **💾 Export Backup** button regularly to save your progress as a JSON file.
Use **📂 Import Backup** to restore from a saved file.

For shared team access to the same data, the natural next step is connecting to a database (Supabase, PlanetScale etc.) — raise an issue if you'd like this feature.

---

## Tech Stack
- Vanilla HTML/CSS/JS (no framework)
- Anthropic Claude API for AI generation
- Chart.js for CBA visualisations
- PptxGenJS for slide export
- SheetJS (XLSX) for Excel export/import
- Mammoth for Word document parsing
- JSZip for PPTX reading
