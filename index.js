const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const Email = require('./models/email'); // Assuming you have this model set up

const app = express();

// Middleware to parse form data
app.use(express.urlencoded({ extended: true }));

// Set EJS as the template engine
app.set('view engine', 'ejs');

// Serve static files from the "public" folder
app.use(express.static('public'));

// MongoDB connection
mongoose.connect('mongodb://localhost:27017/bulkMailer', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

// Setup Multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Nodemailer setup
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER, // Use environment variable
    pass: process.env.EMAIL_PASS,  // Use environment variable
  },
});

// Function to send bulk emails
async function sendBulkEmails(emailList, message, subject, attachment) {
  for (const email of emailList) {
    const mailOptions = {
      from: process.env.EMAIL_USER, // Use environment variable
      to: email,
      subject: subject,
      text: message,
      attachments: attachment ? [attachment] : [], // Add attachment if available
    };

    try {
      const info = await transporter.sendMail(mailOptions);
      console.log(`Email sent to ${email}: ${info.response}`);
    } catch (err) {
      console.error(`Error sending email to ${email}: ${err}`);
    }
  }
}

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('index');
});

// Route to handle file upload and sending emails
app.post('/upload', upload.fields([{ name: 'emailFile' }, { name: 'attachment' }]), async (req, res) => {
  const { subject, message } = req.body;

  // Check for email file upload
  if (!req.files['emailFile'] || req.files['emailFile'].length === 0) {
    return res.status(400).send('No email file uploaded.');
  }

  // Read the uploaded Excel file
  const workbook = xlsx.readFile(req.files['emailFile'][0].path);
  const sheetName = workbook.SheetNames[0];
  const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  console.log('Parsed worksheet:', worksheet); // Log the worksheet for debugging

  // Extract emails from the Excel file
  const emailList = worksheet.map(row => row.Email).filter(email => email);

  console.log('Email list:', emailList); // Log the email list

  if (emailList.length === 0) {
    return res.status(400).send('No valid email addresses found in the uploaded file.');
  }

  // Check if there is an attachment
  let attachment = null;
  if (req.files['attachment']) {
    attachment = {
      filename: req.files['attachment'][0].originalname,
      path: req.files['attachment'][0].path
    };
  }

  // Send emails
  await sendBulkEmails(emailList, message, subject, attachment);

  res.send('Emails sent successfully');
});

// Start the server
const PORT = process.env.PORT || 3000; // Use environment variable for port
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
