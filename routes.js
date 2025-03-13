const express = require("express");
const multer = require("multer");
const path = require("path");
const XLSX = require("xlsx");
const ExcelFile = require("./models/ExcelFile");
const fs = require("fs");

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
router.post("/edit/:fileName", async (req, res) => {
    try {
      const fileName = req.params.fileName;
      const filePath = `./uploads/${fileName}`;
  
      const updatedData = JSON.parse(req.body.updatedData);
  
      // Read the existing workbook
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
  
      // Convert JSON back to worksheet
      const worksheet = XLSX.utils.json_to_sheet(updatedData);
  
      // Update the workbook with the new worksheet
      workbook.Sheets[sheetName] = worksheet;
  
      // Save the updated Excel file
      XLSX.writeFile(workbook, filePath);
  
      res.redirect(`/edit/${fileName}`); // Stay on edit page after saving
    } catch (err) {
      console.error(err);
      res.redirect("/");
    }
  });
  
  // Route to Download Updated Excel File
  router.get("/download/:fileName", (req, res) => {
    const filePath = `./uploads/${req.params.fileName}`;
    res.download(filePath, req.params.fileName);
  });
router.post("/delete/:id", async (req, res) => {
    try {
      const file = await ExcelFile.findById(req.params.id);
      if (!file) return res.redirect("/");
  
      // Delete file from uploads folder
      fs.unlinkSync(`./uploads/${file.fileName}`);
  
      // Delete from database
      await ExcelFile.findByIdAndDelete(req.params.id);
  
      res.redirect("/");
    } catch (err) {
      console.error(err);
      res.redirect("/");
    }
  });

module.exports = router;
