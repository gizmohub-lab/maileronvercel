const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const Email = require('./models/email'); // Ensure this model is correctly set up
require('dotenv').config(); // Load environment variables from .env file

const app = express();

// Middleware to parse form data
app.use(express.urlencoded({ extended: true }));

// Set EJS as the template engine
app.set('view engine', 'ejs');

// Serve static files from the "public" folder
app.set('views', path.join(__dirname, 'views'));

// MongoDB connection using a cloud provider
mongoose.connect(process.env.MONGODB_CONNECT_URI, 
 ).then(() => {
  console.log('Connected to MongoDB');
}).catch(err => {
  console.error('MongoDB connection error:', err);
});

// Use memory storage for multer
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Nodemailer setup
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER, // Use environment variable for email
    pass: process.env.EMAIL_PASS, // Use environment variable for password
  },
});

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('index');
});

// Route to handle file upload and sending emails
app.post('/upload', upload.fields([{ name: 'emailFile' }, { name: 'attachment' }]), async (req, res) => {
  const { subject, message } = req.body;

  // Read the uploaded Excel file from memory
  const workbook = xlsx.read(req.files['emailFile'][0].buffer);
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
      content: req.files['attachment'][0].buffer, // Use buffer for in-memory storage
    };
  }

  // Send emails
  const emailPromises = emailList.map((email) => {
    const mailOptions = {
      from: process.env.EMAIL_USER, // Update to use the actual sender email
      to: email,
      subject: subject,
      text: message,
      attachments: attachment ? [attachment] : [], // Add attachment if available
    };

    return transporter.sendMail(mailOptions).then(info => {
      console.log(`Email sent to ${email}: ${info.response}`);
    }).catch(err => {
      console.error(`Error sending email to ${email}:`, err);
    });
  });

  // Wait for all emails to be sent
  await Promise.all(emailPromises);

  // Send response
  res.send('Emails sent successfully');
});

// Start the server (not used in Serverless functions)
const PORT = process.env.PORT || 3000; // Use environment variable for port
// app.listen(PORT, () => {
//   console.log(`Server is running on port ${PORT}`);
// });

module.exports = app; // Export app for Vercel
