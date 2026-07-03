const fs = require('fs');
const path = require('path');
const { MongoClient } = require('mongodb');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

const MONGODB_URI = process.env.MONGODB_URI;
const JSON_FILE_PATH = path.join(__dirname, '../db.json');

let client = null;
let db = null;
let collection = null;

async function initDB() {
    if (MONGODB_URI) {
        try {
            client = new MongoClient(MONGODB_URI);
            await client.connect();
            db = client.db('bni_dashboard');
            collection = db.collection('state');
            console.log('Connected to MongoDB successfully!');
        } catch (err) {
            console.error('Failed to connect to MongoDB, falling back to JSON file:', err);
            client = null;
        }
    } else {
        console.log('No MONGODB_URI provided. Using local db.json file.');
    }
}

async function getDashboardData() {
    if (client && collection) {
        try {
            const doc = await collection.findOne({ _id: 'main' });
            if (doc) return { data: doc.data || {}, cur: doc.cur || 'overall', remarks: doc.remarks || {} };
        } catch (err) {
            console.error('Error loading from MongoDB:', err);
        }
    }

    // Fallback to local JSON file
    if (fs.existsSync(JSON_FILE_PATH)) {
        try {
            const fileContent = fs.readFileSync(JSON_FILE_PATH, 'utf8');
            return JSON.parse(fileContent);
        } catch (err) {
            console.error('Error reading local db.json:', err);
        }
    }
    return { data: {}, cur: 'overall', remarks: {} };
}

async function saveDashboardData(state) {
    if (client && collection) {
        try {
            await collection.updateOne(
                { _id: 'main' },
                { $set: { data: state.data || {}, cur: state.cur || 'overall', remarks: state.remarks || {} } },
                { upsert: true }
            );
            return;
        } catch (err) {
            console.error('Error saving to MongoDB:', err);
        }
    }

    // Fallback to local JSON file
    try {
        fs.writeFileSync(JSON_FILE_PATH, JSON.stringify(state, null, 2), 'utf8');
    } catch (err) {
        console.error('Error writing local db.json:', err);
    }
}

module.exports = {
    initDB,
    getDashboardData,
    saveDashboardData
};
