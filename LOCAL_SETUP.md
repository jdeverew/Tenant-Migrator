# Running Locally

## Prerequisites

Install these before anything else:

1. **Node.js 20+** — https://nodejs.org/en/download
2. **PostgreSQL 15+** — https://www.postgresql.org/download

---

## Step 1: Download the Code

In Replit, click the three-dot menu (⋮) in the Files panel and choose **Download as ZIP**. Extract the ZIP to a folder on your computer.

---

## Step 2: Create a PostgreSQL Database

Open a terminal and run:

```bash
psql -U postgres
CREATE DATABASE m365migration;
\q
```

---

## Step 3: Create a `.env` File

In the root of the project folder, create a file called `.env` with the following contents:

```env
DATABASE_URL=postgresql://postgres:yourpassword@localhost:5432/m365migration
SESSION_SECRET=any-long-random-string-you-make-up

# Login credentials for the app (defaults to admin / admin if not set)
ADMIN_USERNAME=admin
ADMIN_PASSWORD=admin
```

Replace `yourpassword` with your actual PostgreSQL password.

---

## Step 4: Install Dependencies and Set Up the Database

Open a terminal in the project folder and run these commands one at a time:

```bash
npm install
npm run db:push
```

---

## Step 5: Start the App

```bash
npm run dev
```

The app will be available at: **http://localhost:5000**

---

## Logging In

Use the credentials you set in the `.env` file. If you didn't change them, the defaults are:

- **Username:** `admin`
- **Password:** `admin`

---

## Changing Your Password

You can set a stronger password via the `ADMIN_PASSWORD` environment variable in your `.env` file. Restart the app after changing it.

---

## Environment Variables Reference

| Variable | Required | Description |
|---|---|---|
| `DATABASE_URL` | Yes | PostgreSQL connection string |
| `SESSION_SECRET` | Yes | Secret used to sign session cookies (any long random string) |
| `ADMIN_USERNAME` | No | Login username (default: `admin`) |
| `ADMIN_PASSWORD` | No | Login password (default: `admin`) |
