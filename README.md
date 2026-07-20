# 📖 GHAFES Ministry Reports
### A complete reporting system for Fellowship, CMC, and Zonal reports

---

> **Who is this guide for?**
> This guide is written for someone who has never built or hosted a website before. Every step is explained clearly. If you can follow a recipe, you can do this. Take it one step at a time and don't rush.

---

## 🗂 Table of Contents

1. [What Does This App Do?](#1-what-does-this-app-do)
2. [What You Will Need](#2-what-you-will-need)
3. [Step 1 — Install Node.js on Your Computer](#step-1--install-nodejs-on-your-computer)
4. [Step 2 — Install Git on Your Computer](#step-2--install-git-on-your-computer)
5. [Step 3 — Create a GitHub Account](#step-3--create-a-github-account)
6. [Step 4 — Create a New GitHub Repository](#step-4--create-a-new-github-repository)
7. [Step 5 — Put the Project Files on Your Computer](#step-5--put-the-project-files-on-your-computer)
8. [Step 6 — Test the App on Your Computer](#step-6--test-the-app-on-your-computer)
9. [Step 7 — Upload (Push) the Files to GitHub](#step-7--upload-push-the-files-to-github)
10. [Step 8 — Create a Vercel Account and Deploy](#step-8--create-a-vercel-account-and-deploy)
11. [Step 9 — Set Up Google Sheets Sync (Optional)](#step-9--set-up-google-sheets-sync-optional)
12. [How to Use the App](#how-to-use-the-app)
13. [Making Changes and Updating the Live Site](#making-changes-and-updating-the-live-site)
14. [Troubleshooting — When Things Go Wrong](#troubleshooting--when-things-go-wrong)

---

## 1. What Does This App Do?

This is a **web application** (a website that works like an app) built for GHAFES campus fellowships. It has three main sections:

| Section | Who Uses It | What It Does |
|---|---|---|
| **Fellowship Report** | Campus Presidents | Enter weekly/monthly data for their fellowship |
| **CMC Report** | Campus Ministry Coordinators | See all their fellowships' data in one place, automatically totalled |
| **Zonal Report** | Zonal Coordinators | See every fellowship across the whole zone |

**Extra features:**
- Export any report to **Excel (.xlsx)** with one click
- **Print to PDF** directly from the browser
- **Sync to Google Sheets** automatically when you save
- Works on phones, tablets, and computers
- Two colour themes: **Dark** and **GHAFES Blue (Bright)**

---

## 2. What You Will Need

Before you start, make sure you have:

- ✅ A **computer** (Windows, Mac, or Linux — all work)
- ✅ An **internet connection**
- ✅ A **Google account** (Gmail — for Google Sheets sync)
- ✅ About **30–45 minutes** of uninterrupted time

**Software you will install (all free):**
- Node.js (runs JavaScript on your computer)
- Git (saves and tracks your files)
- A GitHub account (stores your code online for free)
- A Vercel account (publishes your website for free)

---

## Step 1 — Install Node.js on Your Computer

**What is Node.js?**
Think of it like an engine. Your app is written in JavaScript, and Node.js lets your computer run JavaScript outside of a browser. We need it to test the app and build it.

### How to install Node.js:

1. Open your internet browser (Chrome, Edge, Firefox, etc.)
2. Go to: **https://nodejs.org**
3. You will see two big buttons. Click the one that says **"LTS"** (it may show a version number like 20.x.x LTS). LTS means "Long Term Support" — it's the stable, recommended version.
4. A file will download (it will be named something like `node-v20.x.x-...`)
5. Open that downloaded file and follow the installation steps — just keep clicking **Next** and then **Install**. You don't need to change any settings.
6. When it's done, click **Finish**

### Check that it worked:

1. Open the **Terminal** (on Mac: press `Command + Space`, type `Terminal`, press Enter) or **Command Prompt** (on Windows: press the Windows key, type `cmd`, press Enter)
2. Type this exactly and press Enter:
   ```
   node --version
   ```
3. You should see something like `v20.11.0` — any number is fine as long as it appears. If you see an error, restart your computer and try again.
4. Also type this and press Enter:
   ```
   npm --version
   ```
5. You should see a version number (like `10.2.4`). `npm` is a tool that comes with Node.js — it installs extra packages (like the Excel export feature of our app).

---

## Step 2 — Install Git on Your Computer

**What is Git?**
Git is like a "save history" for your code. It tracks every change you make, and it lets you upload your files to GitHub. Think of it like a very powerful "Undo" button that also lets you share your work.

### How to install Git:

**On Windows:**
1. Go to: **https://git-scm.com/download/win**
2. The download should start automatically. Open the file when it downloads.
3. Keep clicking **Next** through all the steps. The default settings are fine. Click **Install** when you see it.
4. Click **Finish** when done.

**On Mac:**
1. Open the Terminal (Command + Space → type Terminal → Enter)
2. Type this and press Enter:
   ```
   git --version
   ```
3. If Git is not installed, your Mac will ask you to install it — click **Install** and follow the instructions.

### Check that it worked:

1. Open Terminal or Command Prompt
2. Type:
   ```
   git --version
   ```
3. You should see something like `git version 2.43.0`. 

---

## Step 3 — Create a GitHub Account

**What is GitHub?**
GitHub is a free website that stores your code. It's like Google Drive but specifically designed for code. Millions of developers use it. Vercel (our hosting service) connects directly to GitHub, which is why we need it.

### How to create an account:

1. Go to: **https://github.com**
2. Click the big green button that says **"Sign up"**
3. Enter your **email address** → click **Continue**
4. Create a **password** → click **Continue**
5. Choose a **username** (this will be in your website's URL, so make it sensible, like `ghafes-se` or your name)
6. Complete the verification puzzle it shows you
7. Click **Create account**
8. Check your email — GitHub will send a code. Enter that code on the website.
9. You will see some setup questions. You can skip them by clicking **"Skip personalization"** at the bottom.
10. You're in! You now have a GitHub account.

---

## Step 4 — Create a New GitHub Repository

**What is a repository (repo)?**
A repository is like a folder on GitHub where all your project files live. "Repository" is just a fancy word for "project folder on GitHub."

### How to create a repository:

1. Make sure you're logged into **https://github.com**
2. Look for the **"+"** icon in the top-right corner of the page → click it
3. Click **"New repository"**
4. Fill in these details:
   - **Repository name:** `ghafes-ministry-reports` (no spaces, use dashes)
   - **Description:** `GHAFES Ministry Reporting System` (optional but helpful)
   - Choose **"Public"** (this is required for the free Vercel plan)
   - ⚠️ **Do NOT** tick "Add a README file" — we already have one
5. Click the green **"Create repository"** button
6. You will see a page with some instructions. **Leave this page open** — you'll need the URL shown on it (it looks like `https://github.com/YOUR-USERNAME/ghafes-ministry-reports.git`)

---

## Step 5 — Put the Project Files on Your Computer

Now you need to put all the project files into a folder on your computer.

### Create the project folder:

1. Decide where you want to keep this project. A good place is your **Desktop** or your **Documents** folder.
2. Create a new folder there and name it exactly: `ghafes-ministry-reports`

### Copy the project files:

Your project should contain these files (you were given them — copy them into your folder):

```
ghafes-ministry-reports/
│
├── src/
│   ├── App.jsx          ← The main app (the big file with all the code)
│   ├── main.jsx         ← Tells React where to start
│   └── index.css        ← Basic styling
│
├── index.html           ← The page that loads in the browser
├── package.json         ← List of packages the app needs
├── vite.config.js       ← Build tool settings
├── vercel.json          ← Vercel deployment settings
├── .gitignore           ← Tells Git which files to ignore
├── google-apps-script.js← The Google script (separate — not used by Vite)
└── README.md            ← This file!
```

> ⚠️ **Important:** Make sure the `src` folder is **inside** the main project folder. The structure must match exactly what is shown above.

---

## Step 6 — Test the App on Your Computer

Before we put it on the internet, let's make sure it works on your computer first.

### Open the project in Terminal:

1. Open Terminal (Mac) or Command Prompt (Windows)
2. You need to "navigate" to your project folder. Type `cd ` (with a space after it), then drag your project folder into the terminal window — this pastes the full path automatically. Then press **Enter**.

   Alternatively, type it manually:
   - **On Mac:** `cd ~/Desktop/ghafes-ministry-reports`
   - **On Windows:** `cd C:\Users\YourName\Desktop\ghafes-ministry-reports`

3. You should see the folder name appear in the terminal prompt.

### Install the app's dependencies:

The app needs some extra packages (like the Excel export tool). Install them by typing:
```
npm install
```
Then press **Enter**.

You will see a lot of text scrolling — this is normal. It's downloading the packages. Wait until it stops (it could take 1–3 minutes depending on your internet speed). You will see something like `added 150 packages` when it's done.

### Start the app:

Type this and press Enter:
```
npm run dev
```

You will see output like:
```
  VITE v5.x.x  ready in 300 ms

  ➜  Local:   http://localhost:5173/
```

### View it in the browser:

1. Open your browser
2. Type `http://localhost:5173` in the address bar and press Enter
3. You should see the GHAFES Ministry Reports app! 🎉

To stop the app, go back to the terminal and press `Ctrl + C`.

---

## Step 7 — Upload (Push) the Files to GitHub

Now we'll connect your folder to GitHub and upload all the files.

### Configure Git with your name (one-time setup):

In your terminal, type these two commands (replace with your real name and email):
```
git config --global user.name "Your Name"
git config --global user.email "your@email.com"
```

### Initialize Git in your project folder:

Make sure you're still in your project folder in the terminal (you should see `ghafes-ministry-reports` in the prompt). Then type these commands **one at a time**, pressing Enter after each:

```
git init
```
(This turns your folder into a Git project)

```
git add .
```
(The dot `.` means "add ALL files" — this tells Git to track everything)

```
git commit -m "First commit — GHAFES Ministry Reports app"
```
(This "saves" your files in Git's history with a message describing what you did)

```
git branch -M main
```
(This makes sure your main branch is called "main")

Now connect to your GitHub repository. Remember the URL from Step 4? Use it here:
```
git remote add origin https://github.com/YOUR-USERNAME/ghafes-ministry-reports.git
```
(Replace `YOUR-USERNAME` with your actual GitHub username)

Now upload everything:
```
git push -u origin main
```

Git will ask for your GitHub username and password.
> ⚠️ **Note:** GitHub no longer accepts regular passwords here. You need a **Personal Access Token** instead. Here's how to get one:
> 1. On GitHub, click your profile picture (top right) → **Settings**
> 2. Scroll down the left sidebar → click **Developer settings**
> 3. Click **Personal access tokens** → **Tokens (classic)**
> 4. Click **Generate new token (classic)**
> 5. Give it a name like "GHAFES Deploy", set expiration to **90 days** or **No expiration**
> 6. Tick the **repo** checkbox (first one in the list)
> 7. Click **Generate token** at the bottom
> 8. **COPY the token immediately** — you won't see it again. Save it somewhere safe (like a notes app).
> 9. Use this token as your "password" when Git asks for it.

After the push completes, go to your GitHub repository page and refresh it — you should see all your files there! ✅

---

## Step 8 — Create a Vercel Account and Deploy

**What is Vercel?**
Vercel is a free service that takes your code from GitHub and makes it into a live website. Every time you update your code and push to GitHub, Vercel automatically updates your live website too.

### Create a Vercel account:

1. Go to: **https://vercel.com**
2. Click **"Sign Up"**
3. Click **"Continue with GitHub"** — this is the easiest option because it connects your accounts automatically
4. Authorise Vercel to access your GitHub account
5. Vercel will ask what kind of work you do — you can just click "Skip" or choose any option

### Import and deploy your project:

1. Once you're in the Vercel dashboard, click the **"Add New…"** button (or **"New Project"**)
2. You will see a list of your GitHub repositories. Find **`ghafes-ministry-reports`** and click **"Import"**
3. Vercel will show you some settings:
   - **Framework Preset:** It should automatically detect **Vite** — if not, select it from the dropdown
   - **Build Command:** `npm run build` (this should be auto-filled)
   - **Output Directory:** `dist` (this should be auto-filled)
   - You don't need to change anything else
4. Click the big **"Deploy"** button
5. Vercel will now build and deploy your app. You'll see a progress screen with logs. This takes about 1–2 minutes.
6. When it's done, you'll see a success screen with a link like:
   **`https://ghafes-ministry-reports.vercel.app`**

7. Click that link — **your app is now live on the internet!** 🎉🎉🎉

### Share the link:

You can share `https://ghafes-ministry-reports.vercel.app` with anyone — they can open it on their phone or computer without installing anything.

> **Note on the free plan:** The free Vercel plan is more than enough for this app. It gives you unlimited deployments and a working HTTPS website.

---

## Step 9 — Set Up Google Sheets Sync (Optional)

This lets the app automatically save report data to a Google Spreadsheet when someone clicks "Save & Sync to Sheets". Very useful for keeping a permanent backup and doing your own analysis.

### Part A — Create the Google Spreadsheet:

1. Go to: **https://sheets.google.com**
2. Click the **"+"** button to create a new blank spreadsheet
3. Give it a name like `GHAFES Ministry Reports 2026`
4. Look at the URL in your browser. It looks like:
   `https://docs.google.com/spreadsheets/d/`**`1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`**`/edit`
5. The bold part in the middle is your **Spreadsheet ID**. Copy it and save it somewhere.

### Part B — Set up the Apps Script:

1. Go to: **https://script.google.com**
2. Click **"New project"**
3. You'll see a text editor with some default code. **Select all of it** (`Ctrl+A` or `Cmd+A`) and **delete it**
4. Open the file `google-apps-script.js` (from your project folder) and **copy everything in it**
5. Paste it into the Google Apps Script editor
6. Find this line near the top:
   ```javascript
   const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
   ```
   Replace `YOUR_SPREADSHEET_ID_HERE` with your actual Spreadsheet ID from Part A. Keep the quotes around it.
7. Click **File → Save** (or `Ctrl+S`). Name the project `GHAFES Sync`.

### Part C — Deploy as a Web App:

1. Click **Deploy** (top right of the Apps Script editor)
2. Click **"New deployment"**
3. Click the **gear icon ⚙** next to "Select type" and choose **"Web app"**
4. Fill in:
   - **Description:** `GHAFES Ministry Reports Sync`
   - **Execute as:** `Me (your@email.com)`
   - **Who has access:** `Anyone`
5. Click **"Deploy"**
6. Google will ask you to authorise — click **"Authorize access"**, choose your Google account, click **"Advanced"** → **"Go to GHAFES Sync (unsafe)"** → **"Allow"**
   > ⚠️ It says "unsafe" because the app is not verified by Google. This is normal for personal Apps Scripts. Your data is safe.
7. Copy the **Web app URL** it gives you (it looks like `https://script.google.com/macros/s/AKfyc...../exec`)

### Part D — Connect the App to Google Sheets:

1. Open your live GHAFES app (the Vercel link)
2. In the top bar, click **⚙ Sheets**
3. Paste your **Web app URL** into the "Apps Script Deployment URL" field
4. Set a sheet name prefix (e.g., `GHAFES Reports`)
5. Click **"Save Settings"**
6. Now when you fill in a fellowship report and click **"Save & Sync to Sheets"**, the data will appear in your Google Spreadsheet automatically! ✅

---

## How to Use the App

### For Campus Presidents (Fellowship Report):

1. Open the app link
2. Make sure you're on the **📝 Fellowship Report** tab
3. Fill in the **Report Header** at the top (fellowship name, your name, president's name, month)
4. Use the **left sidebar** to switch between sections (Witness, Discipleship, Programs, etc.)
5. Use the **"Week 1", "Week 2"** etc. buttons in the sidebar to enter data for each week
6. Click **"Monthly Summary"** to see all your totals calculated automatically
7. Click **"Save & Sync to Sheets"** (or "Save Report") when done
8. Click **"Export Excel"** to download the report as a spreadsheet
9. Click **"Print / PDF"** to print or save as PDF

### For CMC Staff (CMC Report):

1. Click the **📊 CMC Report** tab
2. Fill in your staff info at the top (name, station, month, books reading, devotional life)
3. Select your **Zone** and **CMC Station** — the table automatically shows all your fellowships' data (pulled from submitted fellowship reports)
4. Add your leadership programs, programs activities, and other CMC stats at the bottom
5. Export to Excel or print when done

### For Zonal Coordinators (Zonal Report):

1. Click the **🌐 Zonal Report** tab
2. Select the **Zone** and **Month/Year**
3. See all fellowship data grouped by station, with automatic subtotals
4. Export to Excel for the full zonal breakdown

### Theme toggle:
- Click **☀ Light** or **🌙 Dark** in the top bar to switch between the dark theme and the GHAFES blue (bright) theme

---

## Making Changes and Updating the Live Site

Once your site is live, here's how to update it when you need to make a change:

1. Make your changes to the files on your computer (e.g., edit `src/App.jsx`)
2. Open Terminal, navigate to your project folder
3. Type these three commands in order:
   ```
   git add .
   git commit -m "Describe what you changed here"
   git push
   ```
4. That's it! Vercel automatically detects the push to GitHub and rebuilds your site in about 1–2 minutes. Refresh your Vercel link and the changes will be live.

---

## Troubleshooting — When Things Go Wrong

### ❌ "npm: command not found"
Node.js was not installed correctly. Go back to **Step 1** and reinstall it. Restart your computer after installing.

### ❌ "git: command not found"
Git was not installed correctly. Go back to **Step 2** and reinstall it. Restart your computer after installing.

### ❌ "npm install" shows errors
- Make sure you are in the right folder (you should see `ghafes-ministry-reports` in the terminal prompt)
- Try running `npm install` again — sometimes network issues cause failures on the first try
- Check your internet connection

### ❌ The app opens but shows a blank white screen
- Open your browser's developer console: press `F12` and click the "Console" tab
- Look for red error messages — these tell you what went wrong
- The most common cause is a file in the wrong folder. Double-check the file structure matches Step 5 exactly.

### ❌ "git push" asks for a password and then fails
- Remember: GitHub doesn't accept regular passwords. You need a Personal Access Token.
- Go back to Step 7 and follow the "Personal Access Token" instructions.

### ❌ Vercel build fails
- In your Vercel dashboard, click on your project → click the failed deployment → click "View Build Logs"
- Read the red error messages carefully
- The most common cause: a file is missing or misnamed. Compare your folder to the structure in Step 5.

### ❌ Google Sheets sync doesn't work
- Make sure you clicked **"Allow"** when authorising the Apps Script
- Make sure the Spreadsheet ID in the Apps Script matches your actual spreadsheet
- After any changes to the script, always do a **new deployment** (not "edit existing deployment") — old deployments don't pick up code changes
- Try opening your Web App URL directly in a browser — you should see the text "✓ GHAFES Ministry Reports Sync is active and running."

### ❌ "I can't find my project folder in the terminal"
- Mac: Type `cd ` then drag the folder from Finder into the terminal and press Enter
- Windows: Open File Explorer, navigate to your folder, then click in the address bar and type `cmd` and press Enter — this opens Command Prompt already in that folder

---

## Project File Summary

| File | What It Does |
|---|---|
| `src/App.jsx` | The entire app — all screens, forms, buttons, and logic |
| `src/main.jsx` | Tells React where to start the app on the page |
| `src/index.css` | Global styles (scrollbars, body background, print styles) |
| `index.html` | The HTML page that loads in the browser |
| `package.json` | Lists all packages the app needs (React, SheetJS, Vite) |
| `vite.config.js` | Configuration for the Vite build tool |
| `vercel.json` | Tells Vercel how to serve the app correctly |
| `.gitignore` | Tells Git not to upload `node_modules` and `dist` folders |
| `google-apps-script.js` | Copy this into Google Apps Script for Sheets sync |
| `README.md` | This guide |

---

## Tech Stack (What the App Is Built With)

| Tool | What It Is | Why We Use It |
|---|---|---|
| **React 18** | JavaScript UI library | Powers the interactive interface |
| **Vite 5** | Build tool | Compiles and bundles the app for deployment |
| **SheetJS (xlsx)** | JavaScript library | Generates real Excel (.xlsx) files in the browser |
| **Google Apps Script** | Google's scripting platform | Receives data and writes it to Google Sheets |
| **Vercel** | Hosting platform | Publishes the app to the internet, free |
| **GitHub** | Code storage | Stores the code and triggers Vercel deployments |

---

*Built for GHAFES — Ghana Fellowship of Evangelical Students*
*"For everything that was written in the past was written to teach us." — Romans 15:4*
