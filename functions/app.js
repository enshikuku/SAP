import express from 'express'
import bodyParser from 'body-parser'
import fs from 'fs'
import path from 'path'
import xlsx from 'xlsx'
import dotenv from 'dotenv'

dotenv.config()

const app = express()

app.use(express.static('public'))

app.use(bodyParser.urlencoded({ extended: true }))
app.set('view engine', 'ejs')

app.get('/', (req, res) => {
    res.render('home', { success: false, error: false, message: '', errorMessage: '' })
})

app.get('/register', (req, res) => {
    res.render('register', { error: false })
})

app.post('/register', (req, res) => {
    const data = req.body

    const filePath = path.join(process.cwd(), 'form-data.xlsx')

    let workbook
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath)
    } else {
        workbook = xlsx.utils.book_new()
        workbook.SheetNames.push('Sheet1')
        workbook.Sheets['Sheet1'] = xlsx.utils.aoa_to_sheet([])
    }

    const worksheet = workbook.Sheets['Sheet1']
    const jsonData = xlsx.utils.sheet_to_json(worksheet)

    jsonData.push(data)
    const updatedWorksheet = xlsx.utils.json_to_sheet(jsonData)
    workbook.Sheets['Sheet1'] = updatedWorksheet

    xlsx.writeFile(workbook, filePath)

    res.render('home', { message: 'Your registration is successful! We look forward to seeing you at the event.', success: true, error: false, errorMessage: '' })
})

app.get('/passcode', (req, res) => {
    res.render('passcode', { error: false, errorMessage: '' })
})

app.post('/validate-passcode', (req, res) => {
    const passcode = req.body.passcode
    const correctPasscode = process.env.PASSCODE

    if (passcode === correctPasscode) {
        const filePath = path.join(process.cwd(), 'form-data.xlsx')

        if (!fs.existsSync(filePath)) {
            return res.render('registrations', { registrations: [], error: false, message: '', errorMessage: '' })
        }

        const workbook = xlsx.readFile(filePath)
        const worksheet = workbook.Sheets['Sheet1']
        const registrations = xlsx.utils.sheet_to_json(worksheet)
        res.render('registrations', { registrations, error: false, message: '', errorMessage: '' })
    } else {
        return res.render('passcode', { error: true, errorMessage: "Incorrect passcode. Please try again." })
    }
})

const PORT = 3000
app.listen(PORT, () => {console.log(`Server is running on http://localhost:${PORT}`)})