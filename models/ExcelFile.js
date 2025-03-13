const mongoose = require("mongoose");

const excelFileSchema = new mongoose.Schema({
  heading: String,
  fileName: String,
  uploadedAt: { type: Date, default: Date.now },
});

module.exports = mongoose.model("ExcelFile", excelFileSchema);
