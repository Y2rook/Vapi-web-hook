const express = require("express");
const { google } = require("googleapis");

const app = express();
app.use(express.json({ limit: "5mb" }));

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
  },
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

app.get("/", (req, res) => {
  res.send("Vapi webhook server is running.");
});

app.post("/webhook", async (req, res) => {
  try {
    const body = req.body;
    const message = body.message || body;

    if (message.type !== "end-of-call-report") {
      return res.status(200).send("Ignored (not end-of-call-report)");
    }

    const customer = message.customer || {};
    const analysis = message.analysis || {};

    const name = analysis.structuredData?.name || customer.name || "Unknown";
    const phone = customer.number || "Unknown";
    const email = analysis.structuredData?.email || "Unknown";
    const summary = message.summary || analysis.summary || "No summary";
    const timestamp = new Date().toISOString();

    const row = [timestamp, name, phone, email, summary];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:E`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: [row] },
    });

    console.log("Row added:", row);
    res.status(200).send("OK");
  } catch (err) {
    console.error("Error handling webhook:", err);
    res.status(500).send("Server error");
  }
});

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
