const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(
  cors({
    origin: [
      "http://localhost:8080",
      "http://localhost:8082",
      "http://localhost:3000",
      "http://localhost:5173",
      "https://book-gx-back.vercel.app",
    ],
    credentials: true,
  })
);
app.use(express.json({ limit: "10mb" }));

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({
    status: "OK",
    message: "BookGX Backend is running",
    timestamp: new Date().toISOString(),
  });
});

// Authenticate with Google Sheets using service account
async function createSheetsClient(serviceAccount) {
  try {
    console.log("🔐 Creating authenticated Google Sheets client...");
    console.log("📧 Service Account Email:", serviceAccount.clientEmail);

    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: serviceAccount.clientEmail,
        private_key: serviceAccount.privateKey.replace(/\\n/g, "\n"), // Handle escaped newlines
        project_id: serviceAccount.projectId,
      },
      scopes: [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
      ],
    });

    const sheets = google.sheets({ version: "v4", auth });
    console.log("✅ Google Sheets client created successfully");
    return sheets;
  } catch (error) {
    console.error("❌ Error creating Google Sheets client:", error);
    throw error;
  }
}

// POST /api/fetchSheets - Get all sheet names from a spreadsheet
app.post("/api/fetchSheets", async (req, res) => {
  try {
    const { spreadsheetId, connection } = req.body;

    console.log("🔐 Backend: Fetching sheets with authenticated access...");
    console.log("📄 Spreadsheet ID:", spreadsheetId);
    console.log("🔗 Connection:", connection?.name);

    // Validate request
    if (!spreadsheetId) {
      return res.status(400).json({ error: "spreadsheetId is required" });
    }

    if (
      !connection ||
      !connection.clientEmail ||
      !connection.privateKey ||
      !connection.projectId
    ) {
      return res.status(400).json({
        error:
          "Service account credentials are required (clientEmail, privateKey, projectId)",
      });
    }

    // Create authenticated sheets client
    const sheets = await createSheetsClient({
      clientEmail: connection.clientEmail,
      privateKey: connection.privateKey,
      projectId: connection.projectId,
    });

    // Fetch spreadsheet metadata
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties.title",
    });

    if (response.data.sheets) {
      const sheetNames = response.data.sheets.map(
        (sheet) => sheet.properties.title
      );
      console.log("✅ Successfully fetched sheet names:", sheetNames);

      res.json({
        success: true,
        sheetNames,
        count: sheetNames.length,
        spreadsheetId,
      });
    } else {
      console.warn("⚠️ No sheets found in spreadsheet");
      res.json({
        success: true,
        sheetNames: [],
        count: 0,
        spreadsheetId,
      });
    }
  } catch (error) {
    console.error("❌ Error fetching sheets:", error);

    // Provide detailed error information
    let errorMessage = "Failed to fetch sheets";
    if (error.code === 403) {
      errorMessage =
        "Permission denied - ensure service account has access to the sheet";
    } else if (error.code === 404) {
      errorMessage = "Spreadsheet not found - check the spreadsheet ID";
    } else if (error.message) {
      errorMessage = error.message;
    }

    res.status(500).json({
      success: false,
      error: errorMessage,
      code: error.code || "UNKNOWN_ERROR",
    });
  }
});

// POST /api/fetchHeaders - Get headers from a specific sheet
app.post("/api/fetchHeaders", async (req, res) => {
  try {
    const { spreadsheetId, sheetName, connection, range = "A1:ZZ1" } = req.body;

    console.log("🔐 Backend: Fetching headers with authenticated access...");
    console.log("📄 Spreadsheet ID:", spreadsheetId);
    console.log("📋 Sheet Name:", sheetName);
    console.log("📏 Range:", range);

    // Validate request
    if (!spreadsheetId || !sheetName) {
      return res
        .status(400)
        .json({ error: "spreadsheetId and sheetName are required" });
    }

    if (
      !connection ||
      !connection.clientEmail ||
      !connection.privateKey ||
      !connection.projectId
    ) {
      return res.status(400).json({
        error:
          "Service account credentials are required (clientEmail, privateKey, projectId)",
      });
    }

    // Create authenticated sheets client
    const sheets = await createSheetsClient({
      clientEmail: connection.clientEmail,
      privateKey: connection.privateKey,
      projectId: connection.projectId,
    });

    // Fetch headers
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!${range}`,
      valueRenderOption: "FORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });

    const headers = response.data.values?.[0] || [];
    const filteredHeaders = headers.filter(
      (header) => header && header.trim() !== ""
    );

    console.log("✅ Successfully fetched headers:", filteredHeaders);

    res.json({
      success: true,
      headers: filteredHeaders,
      count: filteredHeaders.length,
      sheetName,
      spreadsheetId,
    });
  } catch (error) {
    console.error("❌ Error fetching headers:", error);

    // Provide detailed error information
    let errorMessage = "Failed to fetch headers";
    if (error.code === 403) {
      errorMessage =
        "Permission denied - ensure service account has access to the sheet";
    } else if (error.code === 404) {
      errorMessage =
        "Sheet not found - check the sheet name and spreadsheet ID";
    } else if (error.code === 400) {
      errorMessage = "Invalid range - check the range format";
    } else if (error.message) {
      errorMessage = error.message;
    }

    res.status(500).json({
      success: false,
      error: errorMessage,
      code: error.code || "UNKNOWN_ERROR",
    });
  }
});

// POST /api/fetchData - Get data from a specific sheet range (bonus endpoint)
app.post("/api/fetchData", async (req, res) => {
  try {
    const { spreadsheetId, sheetName, connection, range } = req.body;

    console.log("🔐 Backend: Fetching data with authenticated access...");
    console.log("📄 Spreadsheet ID:", spreadsheetId);
    console.log("📋 Sheet Name:", sheetName);
    console.log("📏 Range:", range || "Full sheet");

    // Validate request
    if (!spreadsheetId || !sheetName) {
      return res
        .status(400)
        .json({ error: "spreadsheetId and sheetName are required" });
    }

    if (
      !connection ||
      !connection.clientEmail ||
      !connection.privateKey ||
      !connection.projectId
    ) {
      return res.status(400).json({
        error:
          "Service account credentials are required (clientEmail, privateKey, projectId)",
      });
    }

    // Create authenticated sheets client
    const sheets = await createSheetsClient({
      clientEmail: connection.clientEmail,
      privateKey: connection.privateKey,
      projectId: connection.projectId,
    });

    // Fetch data
    const targetRange = range ? `${sheetName}!${range}` : sheetName;
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: targetRange,
      valueRenderOption: "FORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });

    const data = response.data.values || [];

    console.log("✅ Successfully fetched data:", `${data.length} rows`);

    res.json({
      success: true,
      data,
      rowCount: data.length,
      sheetName,
      spreadsheetId,
    });
  } catch (error) {
    console.error("❌ Error fetching data:", error);

    let errorMessage = "Failed to fetch data";
    if (error.code === 403) {
      errorMessage =
        "Permission denied - ensure service account has access to the sheet";
    } else if (error.code === 404) {
      errorMessage =
        "Sheet not found - check the sheet name and spreadsheet ID";
    } else if (error.message) {
      errorMessage = error.message;
    }

    res.status(500).json({
      success: false,
      error: errorMessage,
      code: error.code || "UNKNOWN_ERROR",
    });
  }
});

// POST /api/testAccess - Test access to a spreadsheet
app.post("/api/testAccess", async (req, res) => {
  try {
    const { spreadsheetId, connection } = req.body;

    console.log("🧪 Backend: Testing access to spreadsheet...");
    console.log("📄 Spreadsheet ID:", spreadsheetId);

    // Validate request
    if (!spreadsheetId) {
      return res.status(400).json({ error: "spreadsheetId is required" });
    }

    if (
      !connection ||
      !connection.clientEmail ||
      !connection.privateKey ||
      !connection.projectId
    ) {
      return res.status(400).json({
        error:
          "Service account credentials are required (clientEmail, privateKey, projectId)",
      });
    }

    // Create authenticated sheets client
    const sheets = await createSheetsClient({
      clientEmail: connection.clientEmail,
      privateKey: connection.privateKey,
      projectId: connection.projectId,
    });

    // Test access
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "properties.title",
    });

    const title = response.data.properties?.title || "Unknown";
    console.log("✅ Access test successful:", title);

    res.json({
      success: true,
      hasAccess: true,
      spreadsheetTitle: title,
      spreadsheetId,
    });
  } catch (error) {
    console.error("❌ Access test failed:", error);

    res.json({
      success: true,
      hasAccess: false,
      error: error.message,
      code: error.code || "UNKNOWN_ERROR",
    });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error("🚨 Unhandled error:", error);
  res.status(500).json({
    success: false,
    error: "Internal server error",
    message: error.message,
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: "Endpoint not found",
    availableEndpoints: [
      "GET /health",
      "POST /api/fetchSheets",
      "POST /api/fetchHeaders",
      "POST /api/fetchData",
      "POST /api/testAccess",
    ],
  });
});

// Start server
app.listen(PORT, () => {
  console.log("🚀 BookGX Backend Server Started");
  console.log(`📡 Server running on http://localhost:${PORT}`);
  console.log(`🔐 Ready to handle authenticated Google Sheets requests`);
  console.log(`📋 Available endpoints:`);
  console.log(`   GET  /health`);
  console.log(`   POST /api/fetchSheets`);
  console.log(`   POST /api/fetchHeaders`);
  console.log(`   POST /api/fetchData`);
  console.log(`   POST /api/testAccess`);
});
