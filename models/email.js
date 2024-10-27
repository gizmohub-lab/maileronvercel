const mongoose = require('mongoose');

const emailSchema = new mongoose.Schema({
  email: { type: String, required: true },
  name: String,
  status: { type: String, default: 'pending' },
});

const Email = mongoose.model('Email', emailSchema);

module.exports = Email;
