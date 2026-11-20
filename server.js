// server.js
"use strict";

const express = require("express");
const mysql = require("mysql2/promise");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
const PORT = Number(process.env.PORT) || 3000;

app.disable("x-powered-by");
app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

// Serve frontend files from /public
app.use(express.static(path.join(__dirname, "public")));

const pool = mysql.createPool({
  host: process.env.DB_HOST,
  port: Number(process.env.DB_PORT) || 4000,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,

  // TiDB Cloud Serverless requires secure TLS connection
  ssl: {
    minVersion: "TLSv1.2",
    rejectUnauthorized: true
  },

  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  dateStrings: true
});

// Excel/CSV upload — file is processed in memory.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

const clean = value =>
  value === undefined || value === null ? "" : String(value).trim();

function validId(value) {
  return /^\d+$/.test(String(value || ""));
}

function getErrorMessage(error) {
  return error.sqlMessage || error.message || "Unexpected server error.";
}

function memberResponse(row) {
  return {
    id: Number(row.id),
    name: row.name || "",
    address: row.address || "",
    dob: row.dob || "",
    mobile: row.mobile || "",
    email: row.email || "",
    fatherSpouseName: row.father_spouse_name || "",
    reference: row.reference || "",
    occupation: row.occupation || "",
    nomineeName: row.nominee_name || "",
    nomineeMobile: row.nominee_mobile || "",
    nomineeDob: row.nominee_dob || "",
    nomineeRelationship: row.nominee_relationship || ""
  };
}

function enquiryResponse(row) {
  return {
    id: Number(row.id),
    name: row.name || "",
    mobile: row.mobile || "",
    reference: row.reference || "",
    planMonth: row.plan_month || "",
    planValue: Number(row.plan_value || 0),
    createdAt: row.created_at || ""
  };
}

function dueResponse(row) {
  return {
    id: Number(row.id),
    name: row.name || "",
    mobileNumber: row.mobile_number || "",
    reference: row.reference || "",
    dueAmount: Number(row.due_amount || 0),
    followBy: row.follow_by || "",
    dueDate: row.due_date || "",
    comments: row.comments || ""
  };
}

function validateMember(body = {}) {
  const member = {
    name: clean(body.name),
    address: clean(body.address),
    dob: clean(body.dob) || null,
    mobile: clean(body.mobile),
    email: clean(body.email),
    fatherSpouseName: clean(body.fatherSpouseName),
    reference: clean(body.reference),
    occupation: clean(body.occupation),
    nomineeName: clean(body.nomineeName),
    nomineeMobile: clean(body.nomineeMobile),
    nomineeDob: clean(body.nomineeDob) || null,
    nomineeRelationship: clean(body.nomineeRelationship)
  };

  if (!member.name || !member.mobile) {
    return { error: "Member name and mobile number are required." };
  }

  return { member };
}

function validateEnquiry(body = {}) {
  const enquiry = {
    name: clean(body.name),
    mobile: clean(body.mobile),
    reference: clean(body.reference),
    planMonth: clean(body.planMonth),
    planValue: Number(body.planValue)
  };

  if (
    !enquiry.name ||
    !enquiry.mobile ||
    !enquiry.planMonth ||
    !Number.isFinite(enquiry.planValue) ||
    enquiry.planValue <= 0
  ) {
    return {
      error: "Name, mobile, plan month and valid plan value are required."
    };
  }

  return { enquiry };
}

function validateDue(body = {}) {
  const due = {
    name: clean(body.name),
    mobileNumber: clean(body.mobileNumber),
    reference: clean(body.reference),
    dueAmount: Number(body.dueAmount),
    followBy: clean(body.followBy),
    dueDate: clean(body.dueDate) || null,
    comments: clean(body.comments)
  };

  if (
    !due.name ||
    !Number.isFinite(due.dueAmount) ||
    due.dueAmount < 0
  ) {
    return { error: "Name and a valid due amount are required." };
  }

  return { due };
}

async function ensureTables() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS members (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
      name VARCHAR(255) NOT NULL,
      address TEXT NULL,
      dob DATE NULL,
      mobile VARCHAR(30) NOT NULL,
      email VARCHAR(255) NULL,
      father_spouse_name VARCHAR(255) NULL,
      reference VARCHAR(255) NULL,
      occupation VARCHAR(255) NULL,
      nominee_name VARCHAR(255) NULL,
      nominee_mobile VARCHAR(30) NULL,
      nominee_dob DATE NULL,
      nominee_relationship VARCHAR(100) NULL,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS enquiries (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
      name VARCHAR(255) NOT NULL,
      mobile VARCHAR(30) NOT NULL,
      reference VARCHAR(255) NULL,
      plan_month VARCHAR(30) NOT NULL,
      plan_value DECIMAL(14,2) NOT NULL,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS member_db (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
      name VARCHAR(255) NOT NULL,
      reference VARCHAR(255) NULL,
      joining_date DATE NULL,
      mobile_number VARCHAR(30) NOT NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS due_list (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
      name VARCHAR(255) NOT NULL,
      mobile_number VARCHAR(30) NULL,
      reference VARCHAR(255) NULL,
      due_amount DECIMAL(14,2) NOT NULL DEFAULT 0,
      follow_by VARCHAR(255) NULL,
      due_date DATE NULL,
      comments TEXT NULL,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
  `);
}

// Health check
app.get("/api/health", async (req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ success: true, message: "Server and database connected." });
  } catch (error) {
    console.error("Health check:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// Dashboard totals
app.get("/api/dashboard", async (req, res) => {
  try {
    const [[members]] = await pool.query(
      "SELECT COUNT(*) AS total FROM members"
    );

    const [[enquiries]] = await pool.query(
      "SELECT COUNT(*) AS total FROM enquiries"
    );

    const [[dues]] = await pool.query(`
      SELECT COUNT(*) AS total, COALESCE(SUM(due_amount), 0) AS amount
      FROM due_list
    `);

    res.json({
      totalMembers: Number(members.total),
      totalEnquiries: Number(enquiries.total),
      totalDueList: Number(dues.total),
      totalDueAmount: Number(dues.amount)
    });
  } catch (error) {
    console.error("Dashboard:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// ---------------- MEMBERS ----------------

app.get("/api/members", async (req, res) => {
  try {
    const search = clean(req.query.search);
    let sql = "SELECT * FROM members";
    const params = [];

    if (search) {
      sql += " WHERE name LIKE ? OR mobile LIKE ? OR reference LIKE ?";
      params.push(`%${search}%`, `%${search}%`, `%${search}%`);
    }

    sql += " ORDER BY id DESC";

    const [rows] = await pool.execute(sql, params);
    res.json(rows.map(memberResponse));
  } catch (error) {
    console.error("List members:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.post("/api/members", async (req, res) => {
  const result = validateMember(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const m = result.member;
  let connection;

  try {
    connection = await pool.getConnection();
    await connection.beginTransaction();

    const [insert] = await connection.execute(
      `INSERT INTO members
        (name, address, dob, mobile, email, father_spouse_name, reference,
         occupation, nominee_name, nominee_mobile, nominee_dob, nominee_relationship)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        m.name,
        m.address || null,
        m.dob,
        m.mobile,
        m.email || null,
        m.fatherSpouseName || null,
        m.reference || null,
        m.occupation || null,
        m.nomineeName || null,
        m.nomineeMobile || null,
        m.nomineeDob,
        m.nomineeRelationship || null
      ]
    );

    await connection.execute(
      `INSERT INTO member_db (name, reference, joining_date, mobile_number)
       VALUES (?, ?, CURDATE(), ?)`,
      [m.name, m.reference || null, m.mobile]
    );

    await connection.commit();

    res.status(201).json({
      success: true,
      id: insert.insertId,
      message: "Member added successfully."
    });
  } catch (error) {
    if (connection) await connection.rollback().catch(() => {});
    console.error("Add member:", error);
    res.status(500).json({
      success: false,
      message: "Could not add member.",
      error: getErrorMessage(error)
    });
  } finally {
    if (connection) connection.release();
  }
});

app.get("/api/members/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid member ID."
    });
  }

  try {
    const [rows] = await pool.execute(
      "SELECT * FROM members WHERE id = ?",
      [req.params.id]
    );

    if (!rows.length) {
      return res.status(404).json({
        success: false,
        message: "Member not found."
      });
    }

    res.json(memberResponse(rows[0]));
  } catch (error) {
    console.error("Get member:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.put("/api/members/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid member ID."
    });
  }

  const result = validateMember(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const m = result.member;

  try {
    const [update] = await pool.execute(
      `UPDATE members SET
        name = ?, address = ?, dob = ?, mobile = ?, email = ?,
        father_spouse_name = ?, reference = ?, occupation = ?,
        nominee_name = ?, nominee_mobile = ?, nominee_dob = ?,
        nominee_relationship = ?
       WHERE id = ?`,
      [
        m.name,
        m.address || null,
        m.dob,
        m.mobile,
        m.email || null,
        m.fatherSpouseName || null,
        m.reference || null,
        m.occupation || null,
        m.nomineeName || null,
        m.nomineeMobile || null,
        m.nomineeDob,
        m.nomineeRelationship || null,
        req.params.id
      ]
    );

    if (!update.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Member not found."
      });
    }

    res.json({ success: true, message: "Member updated successfully." });
  } catch (error) {
    console.error("Update member:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.delete("/api/members/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid member ID."
    });
  }

  try {
    const [result] = await pool.execute(
      "DELETE FROM members WHERE id = ?",
      [req.params.id]
    );

    if (!result.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Member not found."
      });
    }

    res.json({ success: true, message: "Member deleted." });
  } catch (error) {
    console.error("Delete member:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// ---------------- MEMBER DB ----------------

app.get("/api/member-db", async (req, res) => {
  try {
    const search = clean(req.query.search);
    let sql = `
      SELECT id, name, reference, joining_date, mobile_number
      FROM member_db
    `;
    const params = [];

    if (search) {
      sql += " WHERE name LIKE ? OR mobile_number LIKE ? OR reference LIKE ?";
      params.push(`%${search}%`, `%${search}%`, `%${search}%`);
    }

    sql += " ORDER BY id DESC";

    const [rows] = await pool.execute(sql, params);

    res.json(
      rows.map(row => ({
        id: Number(row.id),
        name: row.name || "",
        reference: row.reference || "",
        joiningDate: row.joining_date || "",
        mobile: row.mobile_number || ""
      }))
    );
  } catch (error) {
    console.error("Member DB:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// ---------------- ENQUIRIES ----------------

app.get("/api/enquiries", async (req, res) => {
  try {
    const search = clean(req.query.search);
    let sql = "SELECT * FROM enquiries";
    const params = [];

    if (search) {
      sql += " WHERE name LIKE ? OR mobile LIKE ? OR reference LIKE ?";
      params.push(`%${search}%`, `%${search}%`, `%${search}%`);
    }

    sql += " ORDER BY id DESC";

    const [rows] = await pool.execute(sql, params);
    res.json(rows.map(enquiryResponse));
  } catch (error) {
    console.error("List enquiries:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.post("/api/enquiries", async (req, res) => {
  const result = validateEnquiry(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const e = result.enquiry;
  let connection;

  try {
    connection = await pool.getConnection();
    await connection.beginTransaction();

    const [insert] = await connection.execute(
      `INSERT INTO enquiries (name, mobile, reference, plan_month, plan_value)
       VALUES (?, ?, ?, ?, ?)`,
      [e.name, e.mobile, e.reference || null, e.planMonth, e.planValue]
    );

    await connection.execute(
      `INSERT INTO member_db (name, reference, joining_date, mobile_number)
       VALUES (?, ?, CURDATE(), ?)`,
      [e.name, e.reference || null, e.mobile]
    );

    await connection.commit();

    res.status(201).json({
      success: true,
      id: insert.insertId,
      message: "Enquiry saved successfully."
    });
  } catch (error) {
    if (connection) await connection.rollback().catch(() => {});
    console.error("Add enquiry:", error);
    res.status(500).json({
      success: false,
      message: "Could not save enquiry.",
      error: getErrorMessage(error)
    });
  } finally {
    if (connection) connection.release();
  }
});

app.get("/api/enquiries/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid enquiry ID."
    });
  }

  try {
    const [rows] = await pool.execute(
      "SELECT * FROM enquiries WHERE id = ?",
      [req.params.id]
    );

    if (!rows.length) {
      return res.status(404).json({
        success: false,
        message: "Enquiry not found."
      });
    }

    res.json(enquiryResponse(rows[0]));
  } catch (error) {
    console.error("Get enquiry:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.put("/api/enquiries/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid enquiry ID."
    });
  }

  const result = validateEnquiry(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const e = result.enquiry;

  try {
    const [update] = await pool.execute(
      `UPDATE enquiries
       SET name = ?, mobile = ?, reference = ?, plan_month = ?, plan_value = ?
       WHERE id = ?`,
      [
        e.name,
        e.mobile,
        e.reference || null,
        e.planMonth,
        e.planValue,
        req.params.id
      ]
    );

    if (!update.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Enquiry not found."
      });
    }

    res.json({ success: true, message: "Enquiry updated." });
  } catch (error) {
    console.error("Update enquiry:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.delete("/api/enquiries/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid enquiry ID."
    });
  }

  try {
    const [result] = await pool.execute(
      "DELETE FROM enquiries WHERE id = ?",
      [req.params.id]
    );

    if (!result.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Enquiry not found."
      });
    }

    res.json({ success: true, message: "Enquiry deleted." });
  } catch (error) {
    console.error("Delete enquiry:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// ---------------- DUE LIST ----------------

app.get("/api/due-list", async (req, res) => {
  try {
    const search = clean(req.query.search);
    let sql = "SELECT * FROM due_list";
    const params = [];

    if (search) {
      sql += ` WHERE name LIKE ? OR mobile_number LIKE ?
               OR reference LIKE ? OR follow_by LIKE ? OR comments LIKE ?`;
      params.push(
        `%${search}%`,
        `%${search}%`,
        `%${search}%`,
        `%${search}%`,
        `%${search}%`
      );
    }

    sql += " ORDER BY id DESC";

    const [rows] = await pool.execute(sql, params);
    res.json({ rows: rows.map(dueResponse) });
  } catch (error) {
    console.error("List due records:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.post("/api/due-list", async (req, res) => {
  const result = validateDue(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const d = result.due;

  try {
    const [insert] = await pool.execute(
      `INSERT INTO due_list
        (name, mobile_number, reference, due_amount, follow_by, due_date, comments)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [
        d.name,
        d.mobileNumber || null,
        d.reference || null,
        d.dueAmount,
        d.followBy || null,
        d.dueDate,
        d.comments || null
      ]
    );

    res.status(201).json({
      success: true,
      id: insert.insertId,
      message: "Due record added."
    });
  } catch (error) {
    console.error("Add due:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// Keep this route before /api/due-list/:id.
app.delete("/api/due-list/clear", async (req, res) => {
  try {
    await pool.query("DELETE FROM due_list");
    res.json({ success: true, message: "Due List cleared." });
  } catch (error) {
    console.error("Clear due list:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// Excel/CSV import. Headers: Name, Mobile Number, Reference, Due Amount,
// Follow By, Due Date, Comments.
app.post("/api/due-list/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      success: false,
      message: "Please select a file."
    });
  }

  let connection;

  try {
    const workbook = XLSX.read(req.file.buffer, {
      type: "buffer",
      cellDates: false
    });

    const firstSheet = workbook.SheetNames[0];

    if (!firstSheet) {
      return res.status(400).json({
        success: false,
        message: "Excel sheet is empty."
      });
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
      defval: ""
    });

    let insertedRows = 0;
    let skippedRows = 0;

    connection = await pool.getConnection();
    await connection.beginTransaction();

    for (const row of rows) {
      const normalized = {};

      for (const [key, value] of Object.entries(row)) {
        normalized[String(key).trim().toLowerCase().replace(/[_\s-]+/g, "")] =
          value;
      }

      const name = clean(normalized.name);
      const mobileNumber = clean(
        normalized.mobilenumber || normalized.mobile || normalized.phone
      );
      const reference = clean(normalized.reference);

      const rawAmount =
        normalized.dueamount ?? normalized.amount ?? normalized.due;
      const amount = Number(rawAmount);

      const followBy = clean(normalized.followby);
      const comments = clean(normalized.comments);

      let dueDate = clean(normalized.duedate || normalized.followdate);

      if (/^\d+(\.\d+)?$/.test(dueDate)) {
        const parsed = XLSX.SSF.parse_date_code(Number(dueDate));

        if (parsed) {
          dueDate = `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(
            parsed.d
          ).padStart(2, "0")}`;
        }
      }

      if (!name || rawAmount === "" || !Number.isFinite(amount) || amount < 0) {
        skippedRows++;
        continue;
      }

      if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) {
        dueDate = null;
      }

      await connection.execute(
        `INSERT INTO due_list
          (name, mobile_number, reference, due_amount, follow_by, due_date, comments)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [
          name,
          mobileNumber || null,
          reference || null,
          amount,
          followBy || null,
          dueDate || null,
          comments || null
        ]
      );

      insertedRows++;
    }

    await connection.commit();

    res.json({
      success: true,
      insertedRows,
      skippedRows,
      message: `${insertedRows} row(s) imported.`
    });
  } catch (error) {
    if (connection) await connection.rollback().catch(() => {});
    console.error("Due upload:", error);
    res.status(500).json({
      success: false,
      message: "Excel upload failed.",
      error: getErrorMessage(error)
    });
  } finally {
    if (connection) connection.release();
  }
});

app.get("/api/due-list/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid due ID."
    });
  }

  try {
    const [rows] = await pool.execute(
      "SELECT * FROM due_list WHERE id = ?",
      [req.params.id]
    );

    if (!rows.length) {
      return res.status(404).json({
        success: false,
        message: "Due record not found."
      });
    }

    res.json(dueResponse(rows[0]));
  } catch (error) {
    console.error("Get due:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.put("/api/due-list/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid due ID."
    });
  }

  const result = validateDue(req.body);
  if (result.error) {
    return res.status(400).json({ success: false, message: result.error });
  }

  const d = result.due;

  try {
    const [update] = await pool.execute(
      `UPDATE due_list SET
        name = ?, mobile_number = ?, reference = ?, due_amount = ?,
        follow_by = ?, due_date = ?, comments = ?
       WHERE id = ?`,
      [
        d.name,
        d.mobileNumber || null,
        d.reference || null,
        d.dueAmount,
        d.followBy || null,
        d.dueDate,
        d.comments || null,
        req.params.id
      ]
    );

    if (!update.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Due record not found."
      });
    }

    res.json({ success: true, message: "Due record updated." });
  } catch (error) {
    console.error("Update due:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

app.delete("/api/due-list/:id", async (req, res) => {
  if (!validId(req.params.id)) {
    return res.status(400).json({
      success: false,
      message: "Invalid due ID."
    });
  }

  try {
    const [result] = await pool.execute(
      "DELETE FROM due_list WHERE id = ?",
      [req.params.id]
    );

    if (!result.affectedRows) {
      return res.status(404).json({
        success: false,
        message: "Due record not found."
      });
    }

    res.json({ success: true, message: "Due record deleted." });
  } catch (error) {
    console.error("Delete due:", error);
    res.status(500).json({ success: false, error: getErrorMessage(error) });
  }
});

// Unknown API endpoint handler
app.use("/api", (req, res) => {
  res.status(404).json({
    success: false,
    message: "API endpoint not found."
  });
});

// Handle JSON parsing errors and other Express errors.
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    return res.status(400).json({
      success: false,
      message:
        error.code === "LIMIT_FILE_SIZE"
          ? "File size must be 10MB or less."
          : error.message
    });
  }

  if (error instanceof SyntaxError && error.status === 400 && "body" in error) {
    return res.status(400).json({
      success: false,
      message: "Invalid JSON request body."
    });
  }

  console.error("Unhandled error:", error);
  res.status(500).json({
    success: false,
    message: "Internal server error."
  });
});

async function startServer() {
  try {
    console.log("Connecting to database...");
    await pool.query("SELECT 1");
    console.log("Database connected successfully.");

    console.log("Creating database tables...");
    await ensureTables();
    console.log("Database tables ready.");

    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Server running on port ${PORT}`);
    });
  } catch (error) {
    console.error("STARTUP ERROR:", error);
    process.exit(1);
  }
}

startServer();
