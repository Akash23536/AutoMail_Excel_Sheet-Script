const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
require('dotenv').config();

const UserEmailID = process.env.Email_ID;
const UserPassWard = process.env.APP_Passward;


// Load your Excel file
const workbook = xlsx.readFile('./Sheet01.xlsx');
const sheetName = 'Sheet1';
const worksheet = workbook.Sheets[sheetName];
let data = xlsx.utils.sheet_to_json(worksheet);

// Clean up data
data = data.map((row) => {
  // Remove extra quotes and trim whitespace from fields
  for (let key in row) {
    if (typeof row[key] === 'string') {
      row[key] = row[key].replace(/"/g, '').trim();
    }
  }

  // Split multiple email addresses into an array
  if (row.Email) { row.Email = row.Email.split('|').map((email) => email.trim());}

  return row;
});

// Email configuration
const transporter = nodemailer.createTransport({
  pool: true,
  host: "smtp.gmail.com",
  port: 465,
  secure: true,
  auth: {
    user: UserEmailID, // Replace with your Gmail address
    pass: UserPassWard, // Replace with your App Password
  },
});

// Log file setup
const logFile = './email_log.txt';
let emailCount = 0;
fs.writeFileSync(logFile, ''); // Clear previous log

// Function to log email details
const logEmailDetails = (role, company, emails) => {
  const logMessage = `Role: ${role}, Company: ${company}, Emails: ${emails.join(', ')}\n`;
  fs.appendFileSync(logFile, logMessage);
  console.log(logMessage.trim());
};

// Function to send emails
const sendEmail = async (row) => {
  const { Role, Company, Recruiter, ContactNo, Email } = row;

  const mailOptions = {
    from: 'Akash Bhadana <23536akash.2021@gmail.com>',
    to: Email.join(','), // Combine all email addresses into a single string
    subject: `Request for an Interview Opportunity - ${Role}`,
    html: `
      <p>Dear Hiring Manager,</p>
      <p>I hope this message finds you well. I am writing to express my interest in the <b>${Role}</b> role. I believe my experience and skills align closely with the requirements of this position.</p>
      <p>You can reach me at <b>+91 9811290920</b> for further discussion. Attached is my <a href="https://drive.google.com/file/d/1oG1PV2hH2MVtZtJIwDxqUcODiIXUMmu7/view">resume</a> for your reference.</p>
      <p>Thank you for considering my application. I look forward to the opportunity to contribute to your team.</p>
      <p>Best Regards,</p>
      <p><b>Akash Bhadana</b><br>Contact: +91 9811290920</p>`,
  };

  try {
    await transporter.sendMail(mailOptions);
    emailCount++;
    logEmailDetails(Role, Company, Email);
  } catch (error) {
    console.error(`Error sending email to ${Email.join(', ')}:`, error.message);
  }
};


const sendEmailsSynchronously = async () => {
  for (const row of data) {
    if (row.Email && row.Email.length > 0) {
      await sendEmail(row);
      await new Promise((resolve) => setTimeout(resolve, 90000)); // Pause for 1.5 minute
    } else {
      console.log(`No valid email for ${row.Role} at ${row.Company}`);
    }
  }

  fs.appendFileSync(logFile, `\nTotal Emails Sent: ${emailCount}`);
  console.log(`\nAll emails sent. Total: ${emailCount}`);
};


sendEmailsSynchronously();
