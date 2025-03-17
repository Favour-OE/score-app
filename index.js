const express = require("express");
const ExcelJS = require("exceljs");
const app = express();
const port = 3000;

app.use(express.json());
app.use(express.static("."));

let scoresData = {};
let subjectsByClass = {};

app.post("/submit-scores", async (req, res) => {
  const { class: className, name, scores, serial } = req.body;
  if (!scoresData[className]) scoresData[className] = [];

  let newSerial;
  if (serial) {
    // Update existing record
    scoresData[className] = scoresData[className].filter(entry => entry["Serial Number"] !== serial);
    newSerial = serial;
  } else {
    // New record
    newSerial = Math.max(0, ...scoresData[className].map(entry => entry["Serial Number"] || 0)) + 1;
  }

  const data = scores.map(score => ({
    "Class": className,
    "Serial Number": newSerial,
    "Student Name": name,
    "Subject": score.subject,
    "CA": score.ca,
    "Exam": score.exam
  }));
  scoresData[className] = scoresData[className].concat(data);

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Scores");
  const headers = ["Serial Number", "Student Name"].concat(subjectsByClass[className].flatMap(sub => [`${sub} CA`, `${sub} Exam`]));
  worksheet.columns = headers.map(header => ({ header, key: header, width: 15 }));
  
  const studentData = {};
  scoresData[className].forEach(entry => {
    if (!studentData[entry["Serial Number"]]) {
      studentData[entry["Serial Number"]] = { "Serial Number": entry["Serial Number"], "Student Name": entry["Student Name"] };
    }
    studentData[entry["Serial Number"]][`${entry.Subject} CA`] = entry.CA;
    studentData[entry["Serial Number"]][`${entry.Subject} Exam`] = entry.Exam;
  });
  
  Object.values(studentData).forEach(student => worksheet.addRow(student));
  worksheet.getRow(1).font = { bold: true };
  
  try {
    await workbook.xlsx.writeFile(`${className}_scores.xlsx`);
    res.send("Scores submitted successfully!");
  } catch (error) {
    console.error("Excel write error:", error);
    res.status(500).send("Error saving scores!");
  }
});

app.get("/get-scores", (req, res) => {
  const className = req.query.class;
  res.json(scoresData[className] || []);
});

app.get("/check-setup", (req, res) => {
  const className = req.query.class;
  res.json({ setupComplete: !!subjectsByClass[className], subjects: subjectsByClass[className] || [] });
});

app.post("/save-subjects", (req, res) => {
  const { class: className, subjects } = req.body;
  subjectsByClass[className] = subjects;
  res.send("Subjects saved!");
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});