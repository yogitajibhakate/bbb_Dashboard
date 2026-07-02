const express = require('express');
const cors = require('cors');
const path = require('path');
const Groq = require('groq-sdk');
const { initDB, getDashboardData, saveDashboardData } = require('./db');

require('dotenv').config({ path: path.join(__dirname, '../.env') });

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Serve static files from the React frontend build folder
app.use(express.static(path.join(__dirname, '../frontend/dist')));

// Initialize Groq Client
let groq = null;
if (process.env.GROQ_API_KEY) {
    groq = new Groq({ apiKey: process.env.GROQ_API_KEY });
    console.log('Groq SDK initialized successfully.');
} else {
    console.warn('WARNING: GROQ_API_KEY is not set in your .env file. AI parsing will fallback to dummy responses.');
}

// category standard headers mapping
const categoryKeyMap = {
    attendance: 'Attendance Status',
    kyt: 'KYT Done',
    referrals: 'Referrals Given',
    business: 'Business Given',
    inductions: 'Inductions',
    visitors: 'Visitors',
    testimonials: 'Testimonials',
    bbp: 'BBP Score',
    master: 'Member Name'
};

// GET /api/dashboard
app.get('/api/dashboard', async (req, res) => {
    try {
        const state = await getDashboardData();
        res.json(state);
    } catch (err) {
        console.error('Error fetching dashboard:', err);
        res.status(500).json({ error: 'Failed to fetch dashboard data.' });
    }
});

// POST /api/upload-sheet
app.post('/api/upload-sheet', async (req, res) => {
    const { meetingId, fileName, rawRows } = req.body;

    if (!meetingId || !fileName || !rawRows || !rawRows.length) {
        return res.status(400).json({ error: 'Missing meetingId, fileName, or rawRows.' });
    }

    try {
        let parsedSchema = null;

        if (groq) {
            // Send the first 25 rows to Groq LLM for schema detection
            const systemPrompt = `You are a high-precision BNI spreadsheet schema detector.
You will be given a spreadsheet parsed as a 2D JSON array (first 25 rows) along with the file name.
Your task is to identify:
1. The spreadsheet's data category. It must be exactly one of: "master", "attendance", "kyt", "referrals", "business", "inductions", "visitors", "testimonials", "bbp".
2. The 0-based index of the header row (the row containing column titles like "Member Name", "KYT Done", "Business Given", "Referrals", "Present", etc.).
3. The 0-based index of the column containing the members' names (full names of people).
4. The 0-based index of the column containing the metric value for that category. This must be a DIFFERENT column from the name column (unless the category is "master"). Look for the column that contains the score, status (like P/A), amount of business, or counts.

Return ONLY a JSON object in this format:
{
  "category": "category_name",
  "headerRowIndex": number,
  "nameColumnIndex": number,
  "valueColumnIndex": number
}
Return raw JSON ONLY. No explanation, no markdown blocks.`;

            const sampleRows = rawRows.slice(0, 25);
            const userPrompt = `File Name: "${fileName}"\n\nSpreadsheet Data (First 25 rows):\n${JSON.stringify(sampleRows)}`;

            const response = await groq.chat.completions.create({
                model: 'llama-3.3-70b-versatile',
                response_format: { type: 'json_object' },
                messages: [
                    { role: 'system', content: systemPrompt },
                    { role: 'user', content: userPrompt }
                ],
                temperature: 0.1
            });

            const content = response.choices[0].message.content.trim();
            parsedSchema = JSON.parse(content);
            console.log(`Detected schema for ${fileName}:`, parsedSchema);
        } else {
            // Fallback schema detection (Mock)
            console.warn('No Groq key provided. Running basic rule-based detector.');
            const lowerFileName = fileName.toLowerCase();
            let category = 'master';
            if (lowerFileName.includes('attendance')) category = 'attendance';
            else if (lowerFileName.includes('kyt')) category = 'kyt';
            else if (lowerFileName.includes('referral')) category = 'referrals';
            else if (lowerFileName.includes('business') || lowerFileName.includes('revenue')) category = 'business';
            else if (lowerFileName.includes('visitor')) category = 'visitors';
            else if (lowerFileName.includes('testimonial')) category = 'testimonials';
            else if (lowerFileName.includes('bbp')) category = 'bbp';
            else if (lowerFileName.includes('induction')) category = 'inductions';

            let headerRowIndex = 0;
            let nameColumnIndex = 0;
            let valueColumnIndex = 1;

            // 1. Find the header row by looking for a name column
            for (let i = 0; i < Math.min(5, rawRows.length); i++) {
                const nameIdx = rawRows[i].findIndex(c => /name|member|participant/i.test(String(c)));
                if (nameIdx !== -1) {
                    headerRowIndex = i;
                    nameColumnIndex = nameIdx;
                    break;
                }
            }

            // 2. Find the value column index based on the category
            if (category === 'attendance') {
                let bestValIdx = -1;
                for (let col = 0; col < (rawRows[headerRowIndex] || []).length; col++) {
                    if (col === nameColumnIndex) continue;
                    let matchCount = 0;
                    const sampleLimit = Math.min(rawRows.length, headerRowIndex + 15);
                    for (let row = headerRowIndex + 1; row < sampleLimit; row++) {
                        const cellVal = String(rawRows[row][col] || '').trim().toUpperCase();
                        if (/^(P|A|L|S|PRESENT|ABSENT|LATE|SUBSTITUTE)$/.test(cellVal)) {
                            matchCount++;
                        }
                    }
                    if (matchCount > 2) {
                        bestValIdx = col;
                        break;
                    }
                }
                if (bestValIdx !== -1) {
                    valueColumnIndex = bestValIdx;
                } else {
                    const header = rawRows[headerRowIndex] || [];
                    const statusIdx = header.findIndex((c, idx) => idx !== nameColumnIndex && /status|attendance|present/i.test(String(c)));
                    valueColumnIndex = statusIdx !== -1 ? statusIdx : (nameColumnIndex + 1);
                }
            } else {
                const header = rawRows[headerRowIndex] || [];
                const valIdx = header.findIndex((c, idx) => idx !== nameColumnIndex && String(c).trim() !== '' && !/sl\s*no|mobile|email|phone/i.test(String(c)));
                if (valIdx !== -1) {
                    valueColumnIndex = valIdx;
                } else {
                    const fallbackValIdx = header.findIndex((c, idx) => idx !== nameColumnIndex && String(c).trim() !== '');
                    valueColumnIndex = fallbackValIdx !== -1 ? fallbackValIdx : (nameColumnIndex + 1);
                }
            }

            parsedSchema = { category, headerRowIndex, nameColumnIndex, valueColumnIndex };
        }

        // Programmatically loop through all rows in Node.js
        const category = parsedSchema.category || 'master';
        const headerRowIdx = parsedSchema.headerRowIndex !== undefined ? parsedSchema.headerRowIndex : 0;
        const nameColIdx = parsedSchema.nameColumnIndex !== undefined ? parsedSchema.nameColumnIndex : 0;
        const valueColIdx = parsedSchema.valueColumnIndex !== undefined ? parsedSchema.valueColumnIndex : -1;

        const headerRow = rawRows[headerRowIdx] || [];
        const headers = headerRow.map(h => String(h || '').trim());

        // Assign standard metric name to value column to ensure correct keying in database
        if (valueColIdx >= 0 && category && categoryKeyMap[category]) {
            headers[valueColIdx] = categoryKeyMap[category];
        }

        const rows = [];
        for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            if (!row) continue;
            const name = String(row[nameColIdx] || '').trim();

            // Skip empty names, header names, or totals
            if (!name || /total|sum|average|member name|participant/i.test(name.toLowerCase())) {
                continue;
            }

            const rec = { rawName: name };
            headers.forEach((hdr, idx) => {
                if (idx !== nameColIdx && hdr) {
                    rec[hdr] = row[idx] !== undefined ? row[idx] : '';
                }
            });
            rows.push(rec);
        }

        const sheetId = 'f_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
        const newSheet = {
            id: sheetId,
            name: fileName,
            category: category,
            count: rows.length,
            headers: headers,
            rows: rows,
            meeting: meetingId
        };

        const state = await getDashboardData();
        if (!state.data[meetingId]) {
            state.data[meetingId] = [];
        }
        state.data[meetingId].push(newSheet);

        await saveDashboardData(state);
        res.json({ success: true, sheet: newSheet, state });
    } catch (err) {
        console.error('Error parsing/uploading sheet:', err);
        res.status(500).json({ error: 'Failed to process sheet uploading.' });
    }
});

// POST /api/update-sheet-category
app.post('/api/update-sheet-category', async (req, res) => {
    const { meetingId, sheetId, category } = req.body;

    if (!meetingId || !sheetId || !category) {
        return res.status(400).json({ error: 'Missing parameters.' });
    }

    try {
        const state = await getDashboardData();
        if (state.data[meetingId]) {
            const sheet = state.data[meetingId].find(s => s.id === sheetId);
            if (sheet) {
                const oldCategory = sheet.category;
                sheet.category = category;

                // Re-key the row values based on the new category
                const oldMetricKey = categoryKeyMap[oldCategory] || 'Value';
                const newMetricKey = categoryKeyMap[category] || 'Value';

                sheet.headers = ['Member Name', newMetricKey];
                sheet.rows = sheet.rows.map(row => {
                    const val = row[oldMetricKey] !== undefined ? row[oldMetricKey] : '';
                    return {
                        rawName: row.rawName,
                        [newMetricKey]: val
                    };
                });

                await saveDashboardData(state);
                return res.json({ success: true, state });
            }
        }
        res.status(404).json({ error: 'Sheet not found.' });
    } catch (err) {
        console.error('Error updating category:', err);
        res.status(500).json({ error: 'Failed to update sheet category.' });
    }
});

// POST /api/delete-sheet
app.post('/api/delete-sheet', async (req, res) => {
    const { meetingId, sheetId } = req.body;

    if (!meetingId || !sheetId) {
        return res.status(400).json({ error: 'Missing parameters.' });
    }

    try {
        const state = await getDashboardData();
        if (state.data[meetingId]) {
            state.data[meetingId] = state.data[meetingId].filter(s => s.id !== sheetId);
            if (!state.data[meetingId].length) {
                delete state.data[meetingId];
            }
            await saveDashboardData(state);
            return res.json({ success: true, state });
        }
        res.status(404).json({ error: 'Sheet not found.' });
    } catch (err) {
        console.error('Error deleting sheet:', err);
        res.status(500).json({ error: 'Failed to delete sheet.' });
    }
});

// POST /api/update-remark
app.post('/api/update-remark', async (req, res) => {
    const { meetingId, name, remark } = req.body;
    if (!meetingId || !name) {
        return res.status(400).json({ error: 'Missing meetingId or name.' });
    }
    try {
        const state = await getDashboardData();
        if (!state.remarks) {
            state.remarks = {};
        }
        if (!state.remarks[meetingId]) {
            state.remarks[meetingId] = {};
        }
        state.remarks[meetingId][name] = remark;
        await saveDashboardData(state);
        res.json({ success: true, state });
    } catch (err) {
        console.error('Error updating remark:', err);
        res.status(500).json({ error: 'Failed to update remark.' });
    }
});

// POST /api/reset-meeting
app.post('/api/reset-meeting', async (req, res) => {
    const { meetingId } = req.body;
    if (!meetingId) {
        return res.status(400).json({ error: 'Missing meetingId.' });
    }

    try {
        const state = await getDashboardData();
        if (state.data[meetingId]) {
            delete state.data[meetingId];
        }
        if (state.remarks && state.remarks[meetingId]) {
            delete state.remarks[meetingId];
        }
        await saveDashboardData(state);
        res.json({ success: true, state });
    } catch (err) {
        console.error('Error resetting meeting:', err);
        res.status(500).json({ error: 'Failed to reset meeting.' });
    }
});

// POST /api/reset-all
app.post('/api/reset-all', async (req, res) => {
    try {
        const state = { data: {}, cur: 'overall', remarks: {} };
        await saveDashboardData(state);
        res.json({ success: true, state });
    } catch (err) {
        console.error('Error resetting all:', err);
        res.status(500).json({ error: 'Failed to reset all data.' });
    }
});

// For any other GET request, serve index.html (React router support)
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/dist/index.html'));
});

// Start server on port 5050 (with GROQ API Key)
app.listen(PORT, async () => {
    await initDB();
    console.log(`Server is running on port ${PORT}`);
});
