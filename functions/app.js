import express from 'express';
import bodyParser from 'body-parser';
import fs from 'fs';
import path from 'path';
import xlsx from 'xlsx';
import serverless from 'serverless-http';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

// Serve static files from the "public" directory
app.use(express.static('public'));

app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');

// Serve the home page
app.get('/', (req, res) => {
    res.render('home', {success: false});
});

// Serve the registration page
app.get('/register', (req, res) => {
    res.render('register');
});

// Handle form submission
app.post('/register', (req, res) => {
    const data = req.body;

    // Path to the Excel file
    const filePath = path.join(process.cwd(), 'form-data.xlsx');

    // Read the existing Excel file or create a new one
    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
        workbook.SheetNames.push('Sheet1');
        workbook.Sheets['Sheet1'] = xlsx.utils.aoa_to_sheet([]);
    }

    const worksheet = workbook.Sheets['Sheet1'];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    // Append the new data
    jsonData.push(data);
    const updatedWorksheet = xlsx.utils.json_to_sheet(jsonData);
    workbook.Sheets['Sheet1'] = updatedWorksheet;

    // Write the updated Excel file
    xlsx.writeFile(workbook, filePath);

    res.render('home', { message: 'Your registration is successful! We look forward to seeing you at the event.', success: true });
});




if (process.env.NODE_ENV === "dev") {
    const PORT = 3000;
    app.listen(PORT, () => {
        console.log(`Server is running on http://localhost:${PORT}`);
    });
}

// Export the app to be used by Netlify Functions
export const handler = serverless(app)
