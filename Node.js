// server.js
const express = require("express");
const http = require("http");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(express.static(path.join(__dirname, "public")));

// In production you need real certs. For local dev, you can create self-signed certs:
const options = {
  key: fs.readFileSync("localhost.key"),
  cert: fs.readFileSync("localhost.crt"),
};

http.createServer(options, app).listen(3000, () => {
  console.log("Server running at https://localhost:3000");
});
