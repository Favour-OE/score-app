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
  "prenursery": ["maths", "english", "elementary science", "social habits", "health habit", "crk", "writing", "creative art", "craft"],
  "nursery1": ["maths", "english", "elementary science", "social habits", "health habit", "crk", "writing", "creative art", "craft"],
  "nursery2": ["maths", "english", "elementary science", "social habits", "health habit", "crk", "writing", "creative art", "craft"],
  "primary1": ["english language", "mathematics", "elementary science", "social studies", "religion studies", "home economics", "verbal aptitude", "computer studies", "agricultural science", "civic education", "hand writing", "creative art", "quantitative", "health education", "french", "craft"],
  "primary2": ["english language", "mathematics", "elementary science", "social studies", "religion studies", "home economics", "verbal aptitude", "computer studies", "agricultural science", "civic education", "hand writing", "creative art", "quantitative", "health education", "french", "craft"],
  "primary3": ["english language", "mathematics", "elementary science", "social studies", "religion studies", "home economics", "verbal aptitude", "computer studies", "agricultural science", "civic education", "hand writing", "creative art", "quantitative", "health education", "french", "craft"],
  "primary4": ["english language", "mathematics", "elementary science", "social studies", "religion studies", "home economics", "verbal aptitude", "computer studies", "agricultural science", "civic education", "hand writing", "creative art", "quantitative", "health education", "french", "craft"],
  "primary5": ["english language", "mathematics", "elementary science", "social studies", "religion studies", "home economics", "verbal aptitude", "computer studies", "agricultural science", "civic education", "hand writing", "creative art", "quantitative", "health education", "french", "craft"],
  "jss1": ["english language", "mathematics", "intro. technology", "social studies", "civic education", "computer studies", "integrated science", "agric science", "business studies", "physical health education", "french"],
  "jss2": ["english language", "mathematics", "intro. technology", "social studies", "civic education", "computer studies", "integrated science", "agric science", "business studies", "physical health education", "french"],
  "jss3": ["english language", "mathematics", "intro. technology", "social studies", "civic education", "computer studies", "integrated science", "agric science", "business studies", "physical health education", "french"],
  "sss1": ["english language", "mathematics", "christian religious knowledge", "civic", "agric science", "english literature", "biology", "economic", "government", "chemistry", "physic", "french", "data processing", "commerce", "accounting"],
  "sss2a": ["english language", "mathematics", "christian religious knowledge", "civic", "agric science", "english literature", "economic", "government", "french", "data processing"],
  "sss3a": ["english language", "mathematics", "christian religious knowledge", "civic", "agric science", "english literature", "economic", "government", "french", "data processing"],
  "sss2b": ["english language", "mathematics", "civic edu.", "agric science", "biology", "economic", "chemistry", "physics", "french", "data processing", "geography"],
  "sss3b": ["english language", "mathematics", "civic edu.", "agric science", "biology", "economic", "chemistry", "physics", "french", "data processing", "geography"]
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
app.get("/enter-assessment.html", (req, res) => res.sendFile(__dirname + "/enter-assessment.html"));
app.get("/view-records.html", (req, res) => res.sendFile(__dirname + "/view-records.html"));
app.get("/edit-assessment.html", (req, res) => res.sendFile(__dirname + "/edit-assessment.html"));
app.get("/admin.html", (req, res) => res.sendFile(__dirname + "/admin.html"));

app.post("/submit-scores", async (req, res) => {
  const { class: className, name, scores } = req.body;
  const maxSerial = await db.collection("scores").find({ class: className }).sort({ serialNumber: -1 }).limit(1).toArray();
  const newSerial = maxSerial.length > 0 ? maxSerial[0].serialNumber + 1 : 1;

  const data = scores.map(score => ({
    class: className,
    serialNumber: newSerial,
    studentName: name,
    subject: score.subject,
    ca1: score.ca1,
    ca2: score.ca2,
    exam: score.exam,
    total: (score.ca1 || 0) + (score.ca2 || 0) + (score.exam || 0)
  }));

  await db.collection("scores").insertMany(data);
  res.send("Scores submitted");
});

app.post("/update-scores", async (req, res) => {
  const { class: className, serialNumber, scores } = req.body;
  for (let score of scores) {
    await db.collection("scores").updateOne(
      { class: className, serialNumber: parseInt(serialNumber), subject: score.subject },
      { $set: { ca1: score.ca1, ca2: score.ca2, exam: score.exam, total: (score.ca1 || 0) + (score.ca2 || 0) + (score.exam || 0) } },
      { upsert: true }
    );
  }
  res.send("Scores updated");
});

app.get("/get-scores", async (req, res) => {
  const className = req.query.class;
  const rawScores = await db.collection("scores").find({ class: className }).toArray();
  const subjects = (await db.collection("subjects").findOne({ class: className }))?.subjects || [];
  
  res.json({ subjects, records: rawScores });
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

async function generateExcel(className) {
  const scoresData = await db.collection("scores").find({ class: className }).toArray();
  if (scoresData.length === 0) return null;
  const subjects = (await db.collection("subjects").findOne({ class: className }))?.subjects.map(s => s.toUpperCase()) || [];
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Scores");
  const isScienceClass = className === "sss2b" || className === "sss3b";

  const studentData = {};
  scoresData.forEach(entry => {
    if (!studentData[entry.serialNumber]) {
      studentData[entry.serialNumber] = { "S/N": entry.serialNumber, "NAMES": entry.studentName, scores: {} };
    }
    studentData[entry.serialNumber].scores[entry.subject.toUpperCase()] = { ca1: entry.ca1, ca2: entry.ca2, exam: entry.exam, total: entry.total };
  });

  const students = Object.values(studentData);
  const isSenior = className.startsWith("sss");

  for (let student of students) {
    student["GRAND TOTAL"] = 0;
    for (let sub of subjects) {
      const score = student.scores[sub] || { ca1: 0, ca2: 0, exam: 0, total: 0 };
      student[`${sub} CA1`] = score.ca1 === 0 && score.ca2 === 0 && score.exam === 0 ? "-" : score.ca1;
      student[`${sub} CA2`] = score.ca1 === 0 && score.ca2 === 0 && score.exam === 0 ? "-" : score.ca2;
      student[`${sub} EXAM`] = score.ca1 === 0 && score.ca2 === 0 && score.exam === 0 ? "-" : score.exam;
      student[`${sub} TOTAL`] = score.ca1 === 0 && score.ca2 === 0 && score.exam === 0 ? "-" : score.total;
      student["GRAND TOTAL"] += score.total;
    }
    const offeredSubjects = Object.keys(student.scores).length;
    const maxScore = isScienceClass ? 1000 : subjects.length * 100;
    student["GRAND AVERAGE"] = `${((student["GRAND TOTAL"] / maxScore) * 100).toFixed(1)}%`;
  }

  for (let sub of subjects) {
    const subjectTotals = students.map(s => typeof s[`${sub} TOTAL`] === "number" ? s[`${sub} TOTAL`] : 0).sort((a, b) => b - a);
    const offeringStudents = students.filter(s => typeof s[`${sub} TOTAL`] === "number" && s[`${sub} TOTAL`] > 0);
    const classAverage = offeringStudents.length > 0 
      ? (offeringStudents.reduce((sum, s) => sum + s[`${sub} TOTAL`], 0) / (offeringStudents.length * 100) * 100).toFixed(1) + "%" 
      : "-";
    students.forEach(s => {
      const total = typeof s[`${sub} TOTAL`] === "number" ? s[`${sub} TOTAL`] : 0;
      s[`${sub} POSITION`] = total > 0 ? getOrdinal(subjectTotals.indexOf(total) + 1) : "-";
      s[`${sub} GRADE`] = total > 0 ? getGrade(total, isSenior) : "-";
      s[`${sub} CLASS AVERAGE`] = classAverage;
    });
  }

  const grandTotals = students.map(s => s["GRAND TOTAL"]).sort((a, b) => b - a);
  students.forEach(s => {
    s["POSITION (IN CLASS)"] = s["GRAND TOTAL"] > 0 ? getOrdinal(grandTotals.indexOf(s["GRAND TOTAL"]) + 1) : "";
  });

  const headers = ["S/N", "NAMES"].concat(subjects.flatMap(sub => [`${sub} CA1`, `${sub} CA2`, `${sub} EXAM`, `${sub} TOTAL`, `${sub} POSITION`, `${sub} GRADE`, `${sub} CLASS AVERAGE`])).concat(["GRAND TOTAL", "GRAND AVERAGE", "POSITION (IN CLASS)"]);
  worksheet.columns = headers.map(header => ({ header, key: header, width: 15 }));
  students.forEach(student => worksheet.addRow(student));
  worksheet.getRow(1).font = { bold: true };
  return workbook;
}

function getOrdinal(n) {
  if (n === 0) return "";
  const s = ["th", "st", "nd", "rd"];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});