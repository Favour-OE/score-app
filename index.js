require('dotenv').config();
const express = require("express");
const ExcelJS = require("exceljs");
const MongoClient = require("mongodb").MongoClient;
const cron = require("node-cron");
const fetch = require("node-fetch");
const archiver = require("archiver");
const jwt = require("jsonwebtoken"); 
const { body, query, validationResult } = require("express-validator"); // Added express-validator

const app = express();
const port = 3000;
const mongoUri = process.env.MONGO_URI;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD;
const JWT_SECRET = process.env.JWT_SECRET;

if (!mongoUri || !ADMIN_PASSWORD || !JWT_SECRET) {
  console.error("Required environment variables (MONGO_URI, ADMIN_PASSWORD, JWT_SECRET) are missing");
  process.exit(1);
}

const loginAttempts = new Map();
const MAX_ATTEMPTS = 10;
const TIME_WINDOW = 60 * 60 * 1000;

app.use(express.json());
app.use(express.static("public"));

let db;

MongoClient.connect(mongoUri)
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

function rateLimitLogin(req, res, next) {
  const ip = req.ip;
  const username = req.body.username || "unknown";
  const key = `${ip}:${username}`;
  const now = Date.now();
  const attemptData = loginAttempts.get(key) || { count: 0, firstAttemptTime: now };
  
  console.log(`Key: ${key}, Attempts: ${attemptData.count}, Time: ${new Date(attemptData.firstAttemptTime).toISOString()}`);
  
  if (now - attemptData.firstAttemptTime >= TIME_WINDOW) {
    console.log(`Resetting attempts for ${key}`);
    attemptData.count = 0;
    attemptData.firstAttemptTime = now;
  }
  
  if (attemptData.count >= MAX_ATTEMPTS) {
    const timeLeft = Math.ceil((TIME_WINDOW - (now - attemptData.firstAttemptTime)) / (60 * 1000));
    console.log(`Blocking ${key} - Too many attempts. Time left: ${timeLeft} mins`);
    return res.status(429).send(`Too many login attempts for ${username}. Try again in ${timeLeft} minutes.`);
  }
  
  req.attemptData = attemptData;
  req.attemptKey = key;
  next();
}

function verifyToken(req, res, next) {
  const token = req.headers["authorization"]?.split(" ")[1];
  if (!token) return res.status(401).send("Access denied. No token provided.");

  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded;
    next();
  } catch (err) {
    res.status(403).send("Invalid token.");
  }
}

app.get("/ping", (req, res) => res.send("Alive!"));

app.get("/dashboard.html", verifyToken, (req, res) => res.sendFile(__dirname + "/public/dashboard.html"));
app.get("/enter-assessment.html", verifyToken, (req, res) => res.sendFile(__dirname + "/public/enter-assessment.html"));
app.get("/view-records.html", verifyToken, (req, res) => res.sendFile(__dirname + "/public/view-records.html"));
app.get("/edit-assessment.html", verifyToken, (req, res) => res.sendFile(__dirname + "/public/edit-assessment.html"));
app.get("/admin.html", (req, res) => res.sendFile(__dirname + "/public/admin.html"));

app.post("/submit-scores", verifyToken, [
  body("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name"),
  body("name").isString().notEmpty().withMessage("Student name is required"),
  body("scores").isArray({ min: 1 }).withMessage("Scores must be a non-empty array"),
  body("scores.*.subject").isString().notEmpty().withMessage("Subject is required for each score"),
  body("scores.*.ca1").isFloat({ min: 0, max: 20 }).withMessage("CA1 must be a number between 0 and 20"),
  body("scores.*.ca2").isFloat({ min: 0, max: 20 }).withMessage("CA2 must be a number between 0 and 20"),
  body("scores.*.exam").isFloat({ min: 0, max: 60 }).withMessage("Exam must be a number between 0 and 60")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  const { class: className, name, scores } = req.body;
  if (req.user.role === "teacher" && req.user.username !== className) {
    return res.status(403).send("Unauthorized: You can only submit scores for your class.");
  }
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

app.post("/update-scores", verifyToken, [
  body("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name"),
  body("serialNumber").isInt({ min: 1 }).withMessage("Serial number must be a positive integer"),
  body("scores").isArray({ min: 1 }).withMessage("Scores must be a non-empty array"),
  body("scores.*.subject").isString().notEmpty().withMessage("Subject is required for each score"),
  body("scores.*.ca1").isFloat({ min: 0, max: 20 }).withMessage("CA1 must be a number between 0 and 20"),
  body("scores.*.ca2").isFloat({ min: 0, max: 20 }).withMessage("CA2 must be a number between 0 and 20"),
  body("scores.*.exam").isFloat({ min: 0, max: 60 }).withMessage("Exam must be a number between 0 and 60")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  const { class: className, serialNumber, scores } = req.body;
  if (req.user.role === "teacher" && req.user.username !== className) {
    return res.status(403).send("Unauthorized: You can only update scores for your class.");
  }
  for (let score of scores) {
    await db.collection("scores").updateOne(
      { class: className, serialNumber: parseInt(serialNumber), subject: score.subject },
      { $set: { ca1: score.ca1, ca2: score.ca2, exam: score.exam, total: (score.ca1 || 0) + (score.ca2 || 0) + (score.exam || 0) } },
      { upsert: true }
    );
  }
  res.send("Scores updated");
});

app.get("/get-scores", verifyToken, [
  query("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  const className = req.query.class;
  if (req.user.role === "teacher" && req.user.username !== className) {
    return res.status(403).send("Unauthorized: You can only view scores for your class.");
  }
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

app.get("/get-subjects", verifyToken, [
  query("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  const className = req.query.class;
  if (req.user.role === "teacher" && req.user.username !== className) {
    return res.status(403).send("Unauthorized: You can only view subjects for your class.");
  }
  const subjects = await db.collection("subjects").findOne({ class: className });
  res.json(subjects ? subjects.subjects : []);
});

app.get("/download-scores", verifyToken, [
  query("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  const className = req.query.class;
  if (req.user.role === "teacher" && req.user.username !== className) {
    return res.status(403).send("Unauthorized: You can only download scores for your class.");
  }
  const workbook = await generateExcel(className);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=${className}_scores.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

app.post("/admin-login", rateLimitLogin, (req, res) => {
  const { username, password } = req.body;
  const ip = req.ip;
  
  console.log(`Admin login attempt - IP: ${ip}, Username: ${username}, Attempt count: ${req.attemptData.count}`);
  
  if (username === "admin" && password === ADMIN_PASSWORD) {
    loginAttempts.delete(req.attemptKey);
    console.log(`Admin login successful - IP: ${ip}, Attempts reset`);
    const token = jwt.sign({ username, role: "admin" }, JWT_SECRET, { expiresIn: "30m" });
    res.json({ message: "Admin login successful!", token });
  } else {
    req.attemptData.count += 1;
    loginAttempts.set(req.attemptKey, req.attemptData);
    console.log(`Admin login failed - IP: ${ip}, New attempt count: ${req.attemptData.count}`);
    res.status(403).send("Invalid admin credentials!");
  }
});

app.post("/teacher-login", rateLimitLogin, (req, res) => {
  const { username, password } = req.body;
  const ip = req.ip;
  const validClass = username.match(/^(prenursery|nursery[1-2]|primary[1-5]|jss[1-3]|sss[1-3][ab]?)$/);
  
  console.log(`Teacher login attempt - IP: ${ip}, Username: ${username}, Attempt count: ${req.attemptData.count}`);
  
  if (validClass && password === `${username}123`) {
    loginAttempts.delete(req.attemptKey);
    console.log(`Teacher login successful - IP: ${ip}, Attempts reset`);
    const token = jwt.sign({ username, role: "teacher" }, JWT_SECRET, { expiresIn: "30m" });
    res.json({ message: "Teacher login successful!", token });
  } else {
    req.attemptData.count += 1;
    loginAttempts.set(req.attemptKey, req.attemptData);
    console.log(`Teacher login failed - IP: ${ip}, New attempt count: ${req.attemptData.count}`);
    res.status(403).send("Invalid teacher credentials!");
  }
});

app.get("/get-classes", verifyToken, async (req, res) => {
  if (req.user.role !== "admin") return res.status(403).send("Unauthorized: Admin only.");
  const classes = await db.collection("subjects").distinct("class");
  res.json(classes);
});

app.post("/delete-record", verifyToken, [
  body("class").isIn(Object.keys(classSubjects)).withMessage("Invalid class name"),
  body("serialNumbers").isArray({ min: 1 }).withMessage("Serial numbers must be a non-empty array"),
  body("serialNumbers.*").isInt({ min: 1 }).withMessage("Each serial number must be a positive integer")
], async (req, res) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });

  if (req.user.role !== "admin") return res.status(403).send("Unauthorized: Admin only.");
  const { class: className, serialNumbers } = req.body;
  await db.collection("scores").deleteMany({ class: className, serialNumber: { $in: serialNumbers.map(Number) } });
  await db.collection("logs").insertOne({ timestamp: new Date().toISOString(), action: `Deleted records S/N ${serialNumbers.join(", ")} for ${className}` });
  res.send("Records deleted");
});

app.get("/download-all-scores", verifyToken, async (req, res) => {
  if (req.user.role !== "admin") return res.status(403).send("Unauthorized: Admin only.");
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