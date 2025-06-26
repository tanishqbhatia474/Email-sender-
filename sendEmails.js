require('dotenv').config();
const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { exit } = require('process');

// Load Excel file
const workbook = xlsx.readFile('./list.xlsx');
const sheetName = 'Recruiters'; // change if needed
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

// Configure transporter
const newTransporter = () => {
  return nodemailer.createTransport({
    pool: true,
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    },
  });
};

const transporter = newTransporter();

const sendEmail = async (row) => {
  const { Name, Company, Email, Role, Link } = row;
  const firstName = Name.split(' ')[0];

  const mailOptions = {
    from: `Tanishq Bhatia <${process.env.EMAIL_USER}>`,
    to: Email,
    subject: `Application for ${Role} at ${Company}`,
    html: `
<p>Dear ${firstName},</p>

<p>I hope this message finds you well. I am <b>Tanishq Bhatia</b>, a pre-final year student pursuing B.Tech in ECE (major) and CSE (minor) at <b>Delhi Technological University</b>, with a CGPA of <b>8.12</b>.</p>

<p>I am reaching out regarding the <b>${Role}</b> opportunity at <b>${Company}</b>. I have relevant experience through my internships and projects, and I am actively looking for an internship starting <b>15th July 2025</b>.</p>

<p>Here’s a quick overview of my profile:</p>
<ul>
  <li><b>SDE Intern</b> at <b>Avlinq Solutions</b>: Worked on Angular, C# APIs, and improved web performance by 30%</li>
  <li><b>Web Design Intern</b> at <b>DTU (USIP)</b>: Developed Physics Department website for students & faculty</li>
  <li><b>Projects</b>: 
    <ul>
      <li>Travel planner using LLaMA 3 & SerpAPI [<a href="https://github.com/tanishqbhatia474/tourai">TourCraft</a>]</li>
      <li>Food delivery app with Stripe integration [<a href="https://github.com/tanishqbhatia474/food-del-app">Food Delivery App</a>]</li>
    </ul>
  </li>
  <li>Solved <b>1000+ DSA questions</b> on platforms like Leetcode, GFG & Codeforces</li>
  <li><b>1st Place</b> in CodeKaze, Invictus 2025 coding competition (DTU Annual Tech Fest)</li>
  <li><b>Top 130 Teams</b> out of <b>15,000+</b> in HackOn With Amazon Season 5 Hackathon</li>
</ul>

<p>I’ve attached my resume and LinkedIn profile for your reference. ${
    Link ? `You can also view the job opening here: <a href="${Link}">${Role}</a>.` : ""
}</p>

<p>I would love the opportunity to discuss how I can contribute to your team. Thank you for considering my profile.</p>

<p>
Warm regards,<br>
<b>Tanishq Bhatia</b><br>
📧 ${process.env.EMAIL_USER}<br>
📞 +91 8295957676<br>
🔗 <a href="https://www.linkedin.com/in/tanishq-bhatia-371641244/">LinkedIn</a><br>
📄 <a href="https://drive.google.com/file/d/YOUR_RESUME_LINK/view">Resume</a>
</p>
`
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log('✅ Email sent to', Email);
  } catch (error) {
    console.error('❌ Error sending email to', Email, error.message);
  }
};

const sendEmailsSynchronously = async () => {
  for (const row of data) {
    await sendEmail(row);
    await new Promise(resolve => setTimeout(resolve, 10000 + Math.random() * 30000)); // 10–40 sec delay
  }
  console.log("🎉 Done sending all emails.");
  exit();
};

sendEmailsSynchronously();
