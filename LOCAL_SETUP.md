# Running Locally on Windows

## Step 1: Install These Two Programs

1. **Node.js** — https://nodejs.org → click the big **"LTS"** button → install it like any normal program
2. **PostgreSQL** — https://www.postgresql.org/download → click Windows → download and install
   - During install it will ask you to set a password — **write it down**, you need it later

---

## Step 2: Create the Database

After PostgreSQL is installed, open **pgAdmin** from the Start menu, then:

1. Expand **Servers** → **PostgreSQL** in the left panel
2. Right-click **Databases** → **Create** → **Database**
3. Type `m365migration` as the name → click **Save**

---

## Step 3: Create the `.env` File

1. Open your project folder (the unzipped folder)
2. Right-click inside it → **New** → **Text Document**
3. Name it `.env` — make sure it is NOT called `.env.txt`
4. Open it with Notepad and paste this in:

```
DATABASE_URL=postgresql://postgres:YOURPASSWORD@localhost:5432/m365migration
SESSION_SECRET=anylongrandomtextyouwant
ADMIN_USERNAME=admin
ADMIN_PASSWORD=admin
```

5. Replace `YOURPASSWORD` with the password you chose during PostgreSQL install
6. Save the file

---

## Step 4: Install Dependencies

1. Open your project folder in File Explorer
2. Click the address bar at the top → type `cmd` → press Enter
3. In the Command Prompt window that opens, run:

```
npm install
```

Wait for it to finish.

---

## Step 5: Set Up the Database Tables

In the same Command Prompt window, run:

```
dbpush.bat
```

---

## Step 6: Start the App

In the same Command Prompt window, run:

```
start.bat
```

---

## Step 7: Open the App

Go to **http://localhost:5000** in your browser.

Log in with:
- **Username:** `admin`
- **Password:** `admin`

(or whatever you set in the `.env` file)

---

## Stopping the App

Press **Ctrl + C** in the Command Prompt window.

## Starting Again Later

Just open Command Prompt in the project folder and run `start.bat` again.
