const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const Email = require('./models/email');

const app = express();

// Middleware to parse form data
app.use(express.urlencoded({ extended: true }));

// Set EJS as the template engine
app.set('view engine', 'ejs');

// Serve static files from the "public" folder
app.use(express.static('public'));

// MongoDB connection
mongoose.connect('mongodb://localhost:27017/bulkMailer', {
 
});

// Setup Multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Nodemailer setup
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'mailsendergizmo@gmail.com',
    pass: 'tlfl worp cres lwhi',
  },
});

// Function to send bulk emails
function sendBulkEmails(emailList, message, subject) {
  emailList.forEach((email) => {
    const mailOptions = {
      from: 'your-email@gmail.com',
      to: email,
      subject: subject,
      text: message,
    };

    transporter.sendMail(mailOptions, (err, info) => {
      if (err) {
        console.log(`Error: ${err}`);
      } else {
        console.log(`Email sent to ${email}: ${info.response}`);
      }
    });
  });
}

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('index');
});

// Route to handle file upload and sending emails
// Route to handle file upload and sending emails with optional attachment
app.post('/upload', upload.fields([{ name: 'emailFile' }, { name: 'attachment' }]), (req, res) => {
  const { subject, message } = req.body;

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
  emailList.forEach((email) => {
    const mailOptions = {
      from: 'your-email@gmail.com',
      to: email,
      subject: subject,
      text: message,
      attachments: attachment ? [attachment] : [] // Add attachment if available
    };

    transporter.sendMail(mailOptions, (err, info) => {
      if (err) {
        console.log(`Error: ${err}`);
      } else {
        console.log(`Email sent to ${email}: ${info.response}`);
      }
    });
  });

  res.send('Emails sent successfully');
});




// Start the server
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
