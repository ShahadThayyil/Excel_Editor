require('dotenv').config();
const express = require("express");
const mongoose = require("./db");
const routes = require("./routes");

const app = express();
mongoose();
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use(routes);

const PORT = 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
