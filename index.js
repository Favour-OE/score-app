const express = require("express");
const ExcelJS = require("exceljs");
const MongoClient = require("mongodb").MongoClient;
const cron = require("node-cron");
const fetch = require("node-fetch");
const archiver = require("archiver");
const fs = require("fs");
const path = require("path");

const app = express();
const port = 3000;
const mongoUri = "mongodb+srv://scoreappuser:0908@cluster0.sutmk.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "mysecret2023"; // Set in Render dashboard

app.use(express.json());
app.use(express.static("."));

let db;

MongoClient.connect(mongoUri, { useUnifiedTopology: true })
  .then(client => {
    db = client.db("scoreApp");
    console.log("Connected to MongoDB");
  })
  .catch(error => console.error("MongoDB connection error:", error));

// Self-pinging cron job every 10 minutes
cron.schedule("*/10 * * * *", () => {
  fetch("https://my-score-app.onrender.com/ping")
    .then(() => console.log("Pinged to stay awake"))
    .catch(err => console.error("Ping failed:", err));
});

app.get("/ping", (req, res) => res.send("Alive!"));

app.post("/submit-scores", async (req, res) => {
  const { class: className, name, scores, serial } = req.body;
  let newSerial = serial;
  
  if (serial) {
    await db.collection("scores").deleteMany({ class: className, serialNumber: parseInt(serial) });
  } else {
    const maxSerial = await db.collection("scores").find({ class: className }).sort({ serialNumber: -1 }).limit(1).toArray();
    newSerial = maxSerial.length > 0 ? maxSerial[0].serialNumber + 1 : 1;
  }

  const data = scores.map(score => ({
    class: className,
    serialNumber: newSerial,
    studentName: name,
    subject: score.subject,
    ca: score.ca,
    exam: score.exam
  }));

  await db.collection("scores").insertMany(data);
  await generateExcel(className);
  res.send("Scores submitted successfully!");
});

app.get("/get-scores", async (req, res) => {
  const className = req.query.class;
  const scores = await db.collection("scores").find({ class: className }).toArray();
  res.json(scores);
});

app.get("/check-setup", async (req, res) => {
  const className = req.query.class;
  const subjects = await db.collection("subjects").findOne({ class: className });
  res.json({ setupComplete: !!subjects, subjects: subjects ? subjects.subjects : [] });
});

app.post("/save-subjects", async (req, res) => {
  const { class: className, subjects } = req.body;
  await db.collection("subjects").updateOne(
    { class: className },
    { $set: { subjects } },
    { upsert: true }
  );
  res.send("Subjects saved!");
});

app.get("/get-subjects", async (req, res) => {
  const className = req.query.class;
  const subjects = await db.collection("subjects").findOne({ class: className });
  res.json(subjects ? subjects.subjects : []);
});

app.get("/download-scores", async (req, res) => {
  const className = req.query.class;
  const adminPassword = req.query.adminPassword;

  if (adminPassword !== ADMIN_PASSWORD) {
    return res.status(403).send("Invalid admin password!");
  }

  const scores = await db.collection("scores").find({ class: className }).toArray();
  if (scores.length === 0) return res.status(404).send("No scores found for this class!");

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Scores");
  const subjects = (await db.collection("subjects").findOne({ class: className })).subjects;
  const headers = ["Serial Number", "Student Name"].concat(subjects.flatMap(sub => [`${sub} CA`, `${sub} Exam`]));
  worksheet.columns = headers.map(header => ({ header, key: header, width: 15 }));

  const studentData = {};
  scores.forEach(entry => {
    if (!studentData[entry.serialNumber]) {
      studentData[entry.serialNumber] = { "Serial Number": entry.serialNumber, "Student Name": entry.studentName };
    }
    studentData[entry.serialNumber][`${entry.subject} CA`] = entry.ca;
    studentData[entry.serialNumber][`${entry.subject} Exam`] = entry.exam;
  });

  Object.values(studentData).forEach(student => worksheet.addRow(student));
  worksheet.getRow(1).font = { bold: true };

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=${className}_scores.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

app.get("/admin", (req, res) => res.sendFile(path.join(__dirname, "admin.html")));

app.post("/admin-login", (req, res) => {
  const { username, password } = req.body;
  if (username === "admin" && password === ADMIN_PASSWORD) {
    res.send("Admin login successful!");
  } else {
    res.status(403).send("Invalid admin credentials!");
  }
});

app.get("/get-classes", async (req, res) => {
  const classes = await db.collection("subjects").distinct("class");
  res.json(classes);
});

app.post("/delete-record", async (req, res) => {
  const { class: className, serialNumbers } = req.body;
  await db.collection("scores").deleteMany({ class: className, serialNumber: { $in: serialNumbers.map(Number) } });
  await generateExcel(className);
  await db.collection("logs").insertOne({ timestamp: new Date().toISOString(), action: `Deleted records S/N ${serialNumbers.join(", ")} for ${className}` });
  res.send("Records deleted!");
});

app.post("/reset-subjects", async (req, res) => {
  const { class: className, subjects } = req.body;
  await db.collection("subjects").updateOne(
    { class: className },
    { $set: { subjects } },
    { upsert: true }
  );
  await db.collection("logs").insertOne({ timestamp: new Date().toISOString(), action: `Reset subjects for ${className} to ${subjects.join(", ")}` });
  res.send("Subjects reset!");
});

app.get("/download-all-scores", async (req, res) => {
  const classes = await db.collection("subjects").distinct("class");
  const archive = archiver("zip", { zlib: { level: 9 } });
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=all_scores.zip");
  archive.pipe(res);

  for (const className of classes) {
    await generateExcel(className);
    archive.file(`${className}_scores.xlsx`, { name: `${className}_scores.xlsx` });
  }

  archive.finalize();
});

app.get("/get-log", async (req, res) => {
  const logs = await db.collection("logs").find().sort({ timestamp: -1 }).limit(50).toArray();
  res.json(logs);
});

async function generateExcel(className) {
  const scores = await db.collection("scores").find({ class: className }).toArray();
  if (scores.length === 0) return;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Scores");
  const subjects = (await db.collection("subjects").findOne({ class: className })).subjects;
  const headers = ["Serial Number", "Student Name"].concat(subjects.flatMap(sub => [`${sub} CA`, `${sub} Exam`]));
  worksheet.columns = headers.map(header => ({ header, key: header, width: 15 }));

  const studentData = {};
  scores.forEach(entry => {
    if (!studentData[entry.serialNumber]) {
      studentData[entry.serialNumber] = { "Serial Number": entry.serialNumber, "Student Name": entry.studentName };
    }
    studentData[entry.serialNumber][`${entry.subject} CA`] = entry.ca;
    studentData[entry.serialNumber][`${entry.subject} Exam`] = entry.exam;
  });

  Object.values(studentData).forEach(student => worksheet.addRow(student));
  worksheet.getRow(1).font = { bold: true };
  await workbook.xlsx.writeFile(`${className}_scores.xlsx`);
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});