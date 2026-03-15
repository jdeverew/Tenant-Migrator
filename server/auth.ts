import bcrypt from "bcryptjs";
import session from "express-session";
import connectPg from "connect-pg-simple";
import type { Express, RequestHandler } from "express";

const SESSION_TTL = 7 * 24 * 60 * 60 * 1000; // 1 week

export function setupSession(app: Express) {
  const pgStore = connectPg(session);

  const sessionStore = new pgStore({
    conString: process.env.DATABASE_URL,
    createTableIfMissing: true,
    ttl: SESSION_TTL,
    tableName: "sessions",
  });

  app.use(
    session({
      secret: process.env.SESSION_SECRET || "change-me-in-production",
      store: sessionStore,
      resave: false,
      saveUninitialized: false,
      cookie: {
        httpOnly: true,
        secure: false,
        maxAge: SESSION_TTL,
      },
    })
  );
}

export function registerAuthRoutes(app: Express) {
  const adminUsername = process.env.ADMIN_USERNAME || "admin";
  const adminPasswordHash = process.env.ADMIN_PASSWORD_HASH || "";
  const adminPassword = process.env.ADMIN_PASSWORD || "admin";

  app.post("/api/auth/login", async (req, res) => {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({ message: "Username and password are required" });
    }

    if (username !== adminUsername) {
      return res.status(401).json({ message: "Invalid username or password" });
    }

    let valid = false;
    if (adminPasswordHash) {
      valid = await bcrypt.compare(password, adminPasswordHash);
    } else {
      valid = password === adminPassword;
    }

    if (!valid) {
      return res.status(401).json({ message: "Invalid username or password" });
    }

    (req.session as any).user = {
      id: "admin",
      username: adminUsername,
      email: process.env.ADMIN_EMAIL || `${adminUsername}@local`,
    };

    res.json({ message: "Logged in successfully" });
  });

  app.post("/api/auth/logout", (req, res) => {
    req.session.destroy(() => {
      res.json({ message: "Logged out" });
    });
  });

  app.get("/api/auth/user", isAuthenticated, (req, res) => {
    res.json((req.session as any).user);
  });
}

export const isAuthenticated: RequestHandler = (req, res, next) => {
  if ((req.session as any)?.user) {
    return next();
  }
  res.status(401).json({ message: "Unauthorized" });
};
