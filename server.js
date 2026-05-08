// Local development server
// For production deployment, use Vercel (see README.md)

const express = require('express');
const cors = require('cors');
const path = require('path');
const pptxgen = require('pptxgenjs');

const app = express();
const PORT = 3131;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Claude API proxy
app.post('/api/claude', async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: { message: 'API key not set. Run: set ANTHROPIC_API_KEY=sk-ant-...' } });
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify(req.body)
    });
    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: { message: 'Proxy error: ' + err.message } });
  }
});

// Generate slides — reuse the same Vercel function logic
app.post('/api/generate-slides', async (req, res) => {
  try {
    const handler = require('./api/generate-slides.js');
    // Wrap res to match Vercel's API
    const vercelRes = {
      status: (code) => ({ json: (data) => res.status(code).json(data), send: (data) => res.status(code).send(data) }),
      setHeader: (k, v) => res.setHeader(k, v),
      send: (data) => res.send(data),
      json: (data) => res.json(data)
    };
    await handler.default(req, vercelRes);
  } catch(err) {
    console.error('Slides error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

app.listen(PORT, () => {
  console.log('');
  console.log('  Charter OS running at http://localhost:' + PORT);
  console.log('  API key set: ' + (process.env.ANTHROPIC_API_KEY ? 'YES ✓' : 'NO ✗'));
  console.log('');
});
