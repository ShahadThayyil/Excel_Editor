const express = require("express");
const multer = require("multer");
const path = require("path");
const XLSX = require("xlsx");
const ExcelFile = require("./models/ExcelFile");

const router = express.Router();

// Set up Multer for file upload
const storage = multer.diskStorage({
  destination: "./uploads/",
  filename: (req, file, cb) => cb(null, file.fieldname + "-" + Date.now() + path.extname(file.originalname)),
});
const upload = multer({ storage });

// Home Page (Show all uploaded files)
router.get("/", async (req, res) => {
  const files = await ExcelFile.find().sort({ uploadedAt: -1 });
  res.render("index", { files });
});

// Upload Page
router.get("/upload", (req, res) => res.render("upload"));

// Handle File Upload
router.post("/upload", upload.single("excelFile"), async (req, res) => {
  const { heading } = req.body;
  await ExcelFile.create({ heading, fileName: req.file.filename });
  res.redirect("/");
});

// Search by Heading
router.get("/search", async (req, res) => {
  const { query } = req.query;
  const files = await ExcelFile.find({ heading: new RegExp(query, "i") });
  res.render("index", { files });
});

// View Excel File
router.get("/view/:fileName", (req, res) => {
  const filePath = path.join(__dirname, "uploads", req.params.fileName);
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  res.render("view", { data, sheetName });
});

// Edit Page
router.get("/edit/:fileName", (req, res) => {
  const filePath = path.join(__dirname, "uploads", req.params.fileName);
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  res.render("edit", { data, sheetName, fileName: req.params.fileName });
});

// Update Excel File
router.post("/edit/:fileName", (req, res) => {
  const filePath = path.join(__dirname, "uploads", req.params.fileName);
  const newData = JSON.parse(req.body.updatedData);

  const worksheet = XLSX.utils.json_to_sheet(newData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  XLSX.writeFile(workbook, filePath);
  res.redirect("/");
});

module.exports = router;
