const express = require("express");
const ExcelJS = require("exceljs");
const MongoClient = require("mongodb").MongoClient;
const cron = require("node-cron");
const fetch = require("node-fetch");
const archiver = require("archiver");

const app = express();
const port = 3000;
const mongoUri = "mongodb+srv://scoreappuser:0908@cluster0.sutmk.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "admin0908";

app.use(express.json());
app.use(express.static("."));

let db;

MongoClient.connect(mongoUri, { useUnifiedTopology: true })
  .then(client => {
    db = client.db("scoreApp");
    console.log("Connected to MongoDB");
    seedSubjects();
  })
  .catch(error => console.error("MongoDB connection error:", error));

const classSubjects = {
  "prenursery": ["maths", "social habits", "craft", "rhymes", "science", "crk", "health education", "english", "writing", "creative arts"],
  "nursery1": ["maths", "social habits", "craft", "rhymes", "science", "crk", "health education", "english", "writing", "creative arts"],
  "nursery2": ["maths", "social habits", "craft", "rhymes", "science", "crk", "health education", "english", "writing", "creative arts"],
  "primary1": ["maths", "english", "science", "social studies", "agricultural science", "home economics", "health education", "crs", "verbal reasoning", "quantitative reasoning", "writing", "creative arts", "craft", "computer science", "civic", "french"],
  "primary2": ["maths", "english", "science", "social studies", "agricultural science", "home economics", "health education", "crs", "verbal reasoning", "quantitative reasoning", "writing", "creative arts", "craft", "computer science", "civic", "french"],
  "primary3": ["maths", "english", "science", "social studies", "agricultural science", "home economics", "health education", "crs", "verbal reasoning", "quantitative reasoning", "writing", "creative arts", "craft", "computer science", "civic", "french"],
  "primary4": ["maths", "english", "science", "social studies", "agricultural science", "home economics", "health education", "crs", "verbal reasoning", "quantitative reasoning", "writing", "creative arts", "craft", "computer science", "civic", "french"],
  "primary5": ["maths", "english", "science", "social studies", "agricultural science", "home economics", "health education", "crs", "verbal reasoning", "quantitative reasoning", "writing", "creative arts", "craft", "computer science", "civic", "french"],
  "jss1": ["english", "mathematics", "social studies", "civic education", "agricultural science", "integrated science", "computer studies", "phe", "business studies", "introductory technology", "french"],
  "jss2": ["english", "mathematics", "social studies", "civic education", "agricultural science", "integrated science", "computer studies", "phe", "business studies", "introductory technology", "french"],
  "jss3": ["english", "mathematics", "social studies", "civic education", "agricultural science", "integrated science", "computer studies", "phe", "business studies", "introductory technology", "french"],
  "sss1": ["english", "mathematics", "data processing", "literature", "government", "crs", "economics", "agricultural science", "french", "chemistry", "physics", "biology", "geography", "commerce", "accounting", "civic education"],
  "sss2a": ["english", "mathematics", "data processing", "literature", "government", "crs", "economics", "agricultural science", "french", "civic education"],
  "sss3a": ["english", "mathematics", "data processing", "literature", "government", "crs", "economics", "agricultural science", "french", "civic education"],
  "sss2b": ["english", "mathematics", "data processing", "economics", "agricultural science", "chemistry", "physics", "biology", "geography", "civic education"],
  "sss3b": ["english", "mathematics", "data processing", "economics", "agricultural science", "chemistry", "physics", "biology", "geography", "civic education"]
};

async function seedSubjects() {
  for (const [className, subjects] of Object.entries(classSubjects)) {
    await db.collection("subjects").updateOne(
      { class: className },
      { $set: { subjects } },
      { upsert: true }
    );
  }
  console.log("Subjects seeded for all classes");
}

cron.schedule("*/10 * * * *", () => {
  fetch("https://my-score-app.onrender.com/ping")
    .then(() => console.log("Pinged to stay awake"))
    .catch(err => console.error("Ping failed:", err));
});

app.get("/ping", (req, res) => res.send("Alive!"));

app.get("/dashboard.html", (req, res) => res.sendFile(__dirname + "/dashboard.html"));
app.get("/records.html", (req, res) => res.sendFile(__dirname + "/records.html"));
app.get("/edit.html", (req, res) => res.sendFile(__dirname + "/edit.html"));

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
  res.send("Scores submitted");
});

app.post("/update-scores", async (req, res) => {
  const { class: className, serialNumber, scores } = req.body;
  for (let score of scores) {
    await db.collection("scores").updateOne(
      { class: className, serialNumber: parseInt(serialNumber), subject: score.subject },
      { $set: { ca: score.ca, exam: score.exam } },
      { upsert: true }
    );
  }
  res.send("Scores updated");
});

app.get("/get-scores", async (req, res) => {
  const className = req.query.class;
  const rawScores = await db.collection("scores").find({ class: className }).toArray();
  const subjects = (await db.collection("subjects").findOne({ class: className }))?.subjects.map(s => s.toUpperCase()) || [];
  const studentData = {};

  rawScores.forEach(entry => {
    if (!studentData[entry.serialNumber]) {
      studentData[entry.serialNumber] = { serialNumber: entry.serialNumber, name: entry.studentName, scores: {}, grandTotal: 0 };
    }
    studentData[entry.serialNumber].scores[entry.subject] = { ca: entry.ca, exam: entry.exam };
  });

  const students = Object.values(studentData);
  const isSenior = className.startsWith("sss");

  for (let student of students) {
    for (let sub of subjects) {
      const score = student.scores[sub] || { ca: 0, exam: 0 };
      const total = score.ca + score.exam;
      student.scores[sub] = { ...score, total };
      student.grandTotal += total;
    }
    student.grandAverage = `${((student.grandTotal / (subjects.length * 100)) * 100).toFixed(1)}%`;
  }

  for (let sub of subjects) {
    const subjectTotals = students.map(s => s.scores[sub].total).sort((a, b) => b - a);
    const classAverage = (subjectTotals.reduce((sum, total) => sum + total, 0) / (students.length * 100) * 100).toFixed(1) + "%";
    students.forEach(s => {
      const total = s.scores[sub].total;
      s.scores[sub].position = subjectTotals.indexOf(total) + 1 + (subjectTotals.indexOf(total) === subjectTotals.lastIndexOf(total) ? "th" : "st");
      s.scores[sub].grade = getGrade(total, isSenior);
      s.scores[sub].classAverage = classAverage;
    });
  }

  const grandAverages = students.map(s => parseFloat(s.grandAverage)).sort((a, b) => b - a);
  students.forEach(s => {
    s.classPosition = grandAverages.indexOf(parseFloat(s.grandAverage)) + 1 + (grandAverages.indexOf(parseFloat(s.grandAverage)) === grandAverages.lastIndexOf(parseFloat(s.grandAverage)) ? "th" : "st");
  });

  res.json({ subjects, students });
});

function getGrade(total, isSenior) {
  if (isSenior) {
    if (total >= 85) return "A1";
    if (total >= 80) return "B2";
    if (total >= 60) return "C4";
    if (total >= 53) return "C5";
    if (total >= 50) return "C6";
    if (total >= 45) return "D7";
    if (total >= 40) return "E8";
    return "F9";
  } else {
    if (total >= 80) return "A";
    if (total >= 70) return "B";
    if (total >= 60) return "C";
    if (total >= 50) return "P";
    if (total >= 40) return "E";
    return "F";
  }
}

app.get("/get-subjects", async (req, res) => {
  const className = req.query.class;
  const subjects = await db.collection("subjects").findOne({ class: className });
  res.json(subjects ? subjects.subjects : []);
});

app.get("/download-scores", async (req, res) => {
  const className = req.query.class;
  const workbook = await generateExcel(className);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=${className}_scores.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

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
  await db.collection("logs").insertOne({ timestamp: new Date().toISOString(), action: `Deleted records S/N ${serialNumbers.join(", ")} for ${className}` });
  res.send("Records deleted");
});

app.get("/download-all-scores", async (req, res) => {
  const classes = await db.collection("subjects").distinct("class");
  const archive = archiver("zip", { zlib: { level: 9 } });
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=all_scores.zip");
  archive.pipe(res);

  for (const className of classes) {
    const workbook = await generateExcel(className);
    const buffer = await workbook.xlsx.writeBuffer();
    archive.append(buffer, { name: `${className}_scores.xlsx` });
  }

  archive.finalize();
});

app.get("/get-log", async (req, res) => {
  const logs = await db.collection("logs").find().sort({ timestamp: -1 }).limit(50).toArray();
  res.json(logs);
});

async function generateExcel(className) {
  const scoresData = await db.collection("scores").find({ class: className }).toArray();
  if (scoresData.length === 0) return null;
  const subjects = (await db.collection("subjects").findOne({ class: className }))?.subjects.map(s => s.toUpperCase()) || [];
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Scores");

  const studentData = {};
  scoresData.forEach(entry => {
    if (!studentData[entry.serialNumber]) {
      studentData[entry.serialNumber] = { "S/N": entry.serialNumber, "NAMES": entry.studentName, scores: {} };
    }
    studentData[entry.serialNumber].scores[entry.subject] = { ca: entry.ca, exam: entry.exam };
  });

  const students = Object.values(studentData);
  const isSenior = className.startsWith("sss");

  for (let student of students) {
    student["GRAND TOTAL"] = 0;
    for (let sub of subjects) {
      const score = student.scores[sub] || { ca: 0, exam: 0 };
      const total = score.ca + score.exam;
      student[`${sub} CA`] = score.ca;
      student[`${sub} EXAM`] = score.exam;
      student[`${sub} TOTAL`] = total;
      student["GRAND TOTAL"] += total;
    }
    student["GRAND AVERAGE"] = `${((student["GRAND TOTAL"] / (subjects.length * 100)) * 100).toFixed(1)}%`;
  }

  for (let sub of subjects) {
    const subjectTotals = students.map(s => s[`${sub} TOTAL`]).sort((a, b) => b - a);
    const classAverage = (subjectTotals.reduce((sum, total) => sum + total, 0) / (students.length * 100) * 100).toFixed(1) + "%";
    students.forEach(s => {
      const total = s[`${sub} TOTAL`];
      s[`${sub} POSITION`] = subjectTotals.indexOf(total) + 1 + (subjectTotals.indexOf(total) === subjectTotals.lastIndexOf(total) ? "th" : "st");
      s[`${sub} GRADE`] = getGrade(total, isSenior);
      s[`${sub} CLASS AVERAGE`] = classAverage;
    });
  }

  const grandAverages = students.map(s => parseFloat(s["GRAND AVERAGE"])).sort((a, b) => b - a);
  students.forEach(s => {
    s["POSITION (IN CLASS)"] = grandAverages.indexOf(parseFloat(s["GRAND AVERAGE"])) + 1 + (grandAverages.indexOf(parseFloat(s["GRAND AVERAGE"])) === grandAverages.lastIndexOf(parseFloat(s["GRAND AVERAGE"])) ? "th" : "st");
  });

  const headers = ["S/N", "NAMES"].concat(subjects.flatMap(sub => [`${sub} CA`, `${sub} EXAM`, `${sub} TOTAL`, `${sub} POSITION`, `${sub} GRADE`, `${sub} CLASS AVERAGE`])).concat(["GRAND TOTAL", "GRAND AVERAGE", "POSITION (IN CLASS)"]);
  worksheet.columns = headers.map(header => ({ header, key: header, width: 15 }));
  students.forEach(student => worksheet.addRow(student));
  worksheet.getRow(1).font = { bold: true };
  return workbook;
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});