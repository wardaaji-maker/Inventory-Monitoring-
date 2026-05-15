const express = require('express');
const cors = require('cors');
const fs = require('fs');

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json());

const path = require('path');
const DB_FILE = path.join(__dirname, 'db.json');

// Initialize DB if it doesn't exist
if (!fs.existsSync(DB_FILE)) {
    fs.writeFileSync(DB_FILE, JSON.stringify({ inventory: [], categories: [] }, null, 2));
}

// Helper function to read from DB
const readDB = () => {
    const db = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(db);
}

// Helper function to write to DB
const writeDB = (data) => {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
}

// Routes for inventory
app.get('/api/inventory', (req, res) => {
    const db = readDB();
    res.json(db.inventory);
});

app.post('/api/inventory', (req, res) => {
    const db = readDB();
    const newItem = { id: Date.now(), ...req.body };
    db.inventory.push(newItem);
    writeDB(db);
    res.status(201).json(newItem);
});

app.put('/api/inventory/:id', (req, res) => {
    const db = readDB();
    const id = parseInt(req.params.id);
    const index = db.inventory.findIndex(item => item.id === id);
    if (index !== -1) {
        db.inventory[index] = { ...db.inventory[index], ...req.body };
        writeDB(db);
        res.json(db.inventory[index]);
    } else {
        res.status(404).send('Item not found');
    }
});

// Routes for categories
app.get('/api/categories', (req, res) => {
    const db = readDB();
    res.json(db.categories);
});

app.post('/api/categories', (req, res) => {
    const db = readDB();
    const newCategory = { id: Date.now(), ...req.body };
    db.categories.push(newCategory);
    writeDB(db);
    res.status(201).json(newCategory);
});


app.get('/', (req, res) => {
    res.send('Inventory Monitoring API');
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
