require("dotenv").config();
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Load email credentials
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;

// Read recruiter details from Excel
const workbook = XLSX.readFile("recruiters.xlsx");
const sheetName = workbook.SheetNames[0];
const recruiters = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Configure Nodemailer
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: EMAIL_USER,
    pass: EMAIL_PASS,
  },
});

// Email sending function
async function sendEmail(recruiter) {
  const { Name, Email, Company } = recruiter;
  const subject = `Inquiry About Opportunities at ${Company}`;
  
  const body = `
    Dear ${Name},<br><br>
    
    I hope this email finds you well. My name is Hitesh Soneta, and I am currently pursuing my Master’s in Computer Software Engineering at Northeastern University (Boston campus), set to graduate in August 2025 with a GPA of 3.6. With around four years of professional experience in software engineering, I am eager to contribute my expertise to your organization.<br><br>
    
    Over the course of my career, I’ve developed and deployed full-stack applications, enhanced database performance, and implemented data visualization solutions to drive business insights. I’ve also worked on projects involving financial data visualization, API integrations, and process automation, which have significantly improved workflows and user experiences.<br><br>
    
    I would greatly appreciate it if you could consider me for suitable roles within your organization. I’ve attached my resume for your reference and would be happy to provide additional information or discuss how my skills and experience could be of value to your organization.<br><br>
    
    Thank you for your time and consideration. I look forward to the opportunity to connect and explore potential opportunities.<br><br>
    
    Best regards,<br>
    Hitesh Soneta<br>
    Boston, MA<br>
    +1 (857)-707-6242<br>
    https://www.linkedin.com/in/hitesh-soneta-863a56186
  `;

  const mailOptions = {
    from: EMAIL_USER,
    to: Email,
    subject: subject,
    html: body,
    attachments: [
      {
        filename: "Hitesh Soneta.pdf",
        path: path.join(__dirname, "Hitesh Soneta.pdf"),
      },
    ],
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${Email} (${Company})`);
  } catch (error) {
    console.error(`Failed to send email to ${Email}:`, error.message);
  }
}

// Send emails to all recruiters
async function sendAllEmails() {
  for (let recruiter of recruiters) {
    await sendEmail(recruiter);
  }
}

sendAllEmails();
