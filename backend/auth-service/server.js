require("dotenv").config();
const express = require("express");
const session = require("express-session");
const mysql = require("mysql2/promise");
const bcrypt = require("bcrypt");
const path = require("path");
const app = express();
const PORT = process.env.PORT;
function log(message) {
    console.log(`[${new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" })}] ${message}`);
}
const dbConfig = {
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME
};
const pool = mysql.createPool({
    ...dbConfig,
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.use("/logo", express.static(path.join(__dirname, "..", "NIMIT LOGO.png")));
app.use(express.static(path.join(__dirname, "../../frontend/auth-service")));
app.use(session({
    name: "login_session",
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    rolling: true,
    cookie: {
        httpOnly: true,
        secure: false,
        sameSite: "lax",
        maxAge: 60 * 60 * 1000
    }
}));

app.use(express.static(path.join(__dirname, "../../frontend/auth-service")));

const roleUrls = {
    ADMIN: process.env.ADMIN_URL,
    PROCUREMENT_MANAGER: process.env.PROCUREMENT_MANAGER_URL,
    PROCUREMENT: process.env.PROCUREMENT_URL
};

function getRoleUrl(role) {
    return roleUrls[role] || null;
}

app.get("/", (req, res) => {
    log("GET / - Checking session");

    if (!req.session.user) {
        log("No active session - Serving login page");
        return res.sendFile(path.join(__dirname, "../../frontend/auth-service/index.html"));
    }

    log(`Active session found - User: ${req.session.user.username}, Role: ${req.session.user.role}`);

    const url = getRoleUrl(req.session.user.role);

    if (!url) {
        log(`ERROR - URL not configured for role: ${req.session.user.role}`);
        return res.status(500).send("Role URL is not configured");
    }

    log(`Redirecting ${req.session.user.username} to ${url}`);
    return res.redirect(url);
});

app.post("/login", async (req, res) => {
    try {
        const { username, password } = req.body;

        log(`POST /login - Login attempt for username: ${username}`);

        if (!username || !password) {
            log("Login failed - Username or password missing");
            return res.status(400).json({ success: false, message: "Username and password are required" });
        }

        log(`Checking database for user: ${username}`);

        const [rows] = await pool.execute(
            "SELECT user_id, username, password_hash, role, is_active FROM users WHERE username = ? LIMIT 1",
            [username]
        );

        if (rows.length === 0) {
            log(`Login failed - User not found: ${username}`);
            return res.status(401).json({ success: false, message: "Invalid username or password" });
        }

        const user = rows[0];

        log(`User found - ID: ${user.user_id}, Role: ${user.role}`);

        if (!user.is_active) {
            log(`Login failed - User has no access: ${username}`);

            await pool.execute(
                "INSERT INTO login_logs (user_id, login_status) VALUES (?, 'FAILED')",
                [user.user_id]
            );

            log(`FAILED login recorded for user ID: ${user.user_id}`);

            return res.status(403).json({ success: false, message: "User account has no access" });
        }

        log(`Checking bcrypt password for user: ${username}`);

        const passwordCorrect = await bcrypt.compare(password, user.password_hash);

        if (!passwordCorrect) {
            log(`Login failed - Incorrect password for: ${username}`);

            await pool.execute(
                "INSERT INTO login_logs (user_id, login_status) VALUES (?, 'FAILED')",
                [user.user_id]
            );

            log(`FAILED login recorded for user ID: ${user.user_id}`);

            return res.status(401).json({ success: false, message: "Invalid username or password" });
        }

        log(`Password verified successfully for: ${username}`);

        await pool.execute(
            "INSERT INTO login_logs (user_id, login_status) VALUES (?, 'SUCCESS')",
            [user.user_id]
        );

        log(`SUCCESS login recorded for user ID: ${user.user_id}`);

        req.session.user = {
            user_id: user.user_id,
            username: user.username,
            role: user.role
        };

        req.session.save(err => {
            if (err) {
                log(`ERROR - Failed to save session for: ${username}`);
                return res.status(500).json({ success: false, message: "Failed to create session" });
            }

            const url = getRoleUrl(user.role);

            if (!url) {
                log(`ERROR - URL not configured for role: ${user.role}`);
                return res.status(500).json({ success: false, message: "Role URL is not configured" });
            }

            log(`Login completed - Redirecting ${username} to ${url}`);

            return res.json({
                success: true,
                role: user.role,
                redirect_url: url
            });
        });
    } catch (error) {
        log(`ERROR during login: ${error.message}`);
        res.status(500).json({ success: false, message: "Internal server error" });
    }
});

app.get("/verify", (req, res) => {
    log("GET /verify - Checking authentication");

    if (!req.session.user) {
        log("Verification failed - No active session");
        return res.status(401).json({ authenticated: false });
    }

    log(`Verification successful - User: ${req.session.user.username}, Role: ${req.session.user.role}`);

    return res.json({
        authenticated: true,
        user_id: req.session.user.user_id,
        username: req.session.user.username,
        role: req.session.user.role
    });
});

app.post("/logout", (req, res) => {
    log("POST /logout - Logout requested");

    if (!req.session) {
        log("No session found - Logout completed");
        return res.json({ success: true });
    }

    const username = req.session.user?.username || "Unknown user";

    req.session.destroy(err => {
        if (err) {
            log(`ERROR - Failed to destroy session for: ${username}`);
            return res.status(500).json({ success: false, message: "Failed to logout" });
        }

        res.clearCookie("login_session");
        log(`Logout successful - Session removed from RAM for: ${username}`);

        return res.json({ success: true });
    });
});

app.listen(PORT, async () => {
    log(`Starting login service on port ${PORT}`);

    try {
        await pool.query("SELECT 1");
        log("Database connection successful");
        log(`Login service running at http://localhost:${PORT}`);
    } catch (error) {
        log(`Database connection failed: ${error.message}`);
    }
});