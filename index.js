import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import { appendReport } from './sheetsService.js';

const app = express();

app.use(cors());
app.use(bodyParser.json());

// Health check endpoint to verify server is up
app.get('/ping', (req, res) => {
  res.status(200).send('Sheet API is running');
});

// Main POST endpoint to receive report data
app.post('/report', async (req, res) => {
  try {
    const { merchandiser, outlet, date, notes, items } = req.body;

    if (!merchandiser || !outlet || !date || !Array.isArray(items)) {
      return res.status(400).json({ error: '❌ Invalid payload format' });
    }

    // Pass the full items array and notes to appendReport
    await appendReport(merchandiser, outlet, date, notes, items);

    return res.json({ status: '✅ Report appended to Google Sheet' });
  } catch (error) {
    console.error('❌ Error in /report:', error.message);
    return res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 8080;

app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});