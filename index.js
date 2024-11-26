const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());

// Path to the Excel file
const filePath = path.join(__dirname, 'dummy_posts.xlsx');

// Helper function to read data from Excel
function readExcelFile(filePath) {
    if (!fs.existsSync(filePath)) return [];
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
}

// Helper function to write data to Excel
function writeExcelFile(filePath, data) {
    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');
    xlsx.writeFile(workbook, filePath);
}

// CREATE: Add new data (POST)
app.post('/posts', (req, res) => {
    try {
        const newData = req.body;
        const data = readExcelFile(filePath);
        data.push(newData); // Add new data
        writeExcelFile(filePath, data); // Save to Excel
        res.status(201).json({
            success: true,
            message: 'Data added successfully!',
            data: newData
        });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// READ: Get all data (GET)
app.get('/posts', (req, res) => {
    try {
        const data = readExcelFile(filePath);
        res.json({
            success: true,
            data: data
        });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// UPDATE: Modify data by Post ID (PUT)
app.put('/posts/:id', (req, res) => {
    try {
        const postId = parseInt(req.params.id, 10);
        const updatedData = req.body;
        const data = readExcelFile(filePath);
        const index = data.findIndex((item) => item['Post ID'] === postId);

        if (index === -1) {
            return res.status(404).json({
                success: false,
                message: 'Post not found!'
            });
        }

        data[index] = { ...data[index], ...updatedData }; // Update the record
        writeExcelFile(filePath, data); // Save to Excel

        res.json({
            success: true,
            message: 'Data updated successfully!',
            data: data[index]
        });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// DELETE: Remove data by Post ID (DELETE)
app.delete('/posts/:id', (req, res) => {
    try {
        const postId = parseInt(req.params.id, 10);
        let data = readExcelFile(filePath);
        const initialLength = data.length;

        data = data.filter((item) => item['Post ID'] !== postId);

        if (data.length === initialLength) {
            return res.status(404).json({
                success: false,
                message: 'Post not found!'
            });
        }

        writeExcelFile(filePath, data); // Save to Excel

        res.json({
            success: true,
            message: 'Data deleted successfully!'
        });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
