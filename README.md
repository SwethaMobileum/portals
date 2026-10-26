# Mobileum Caller ID Platform · Philippines (+63)

A professional, enterprise-grade Caller ID verification and compliance demo suite for the Philippines market.

The platform provides a synchronized ecosystem where **Telecom Operators**, **Enterprise Brands (Banks & Corporations)**, and **Customer Smartphone Subscribers** interact seamlessly through a **single shared Excel database workbook** (`shared/CallerID-Database.xlsx`).

---

## 🌐 Portals Overview

| Portal | Local Port | Cloud / Vercel Path | Description |
|---|---|---|---|
| **Portal Hub (Launcher)** | `http://localhost:3000` | `/` | Central launchpad with real-time Excel sync status and portal switchers |
| **Operator Console** | `http://localhost:4001` | `/operator` | Telecom admin: publish number packages, triage & review enterprise applications (Approve/Reject), audit database |
| **Enterprise Portal** | `http://localhost:4002` | `/enterprise` | Corporate self-service: 3-step registration wizard (Package &rarr; Brand & Logo &rarr; Departments & Authorized Numbers) |
| **Customer Phone** | `http://localhost:4003` | `/customer` | Interactive Philippine smartphone simulator: dialpad, dynamic island alert, verified caller badge preview |
| **Live Excel Sheet** | `http://localhost:4001/sheet.html` | `/sheet.html` | Browser-based interactive spreadsheet replica with live auto-refreshing cells and media thumbnails |

---

## ⚡ Tech Stack & Architecture

- **Backend**: Node.js & Express with multipart form file upload support (`multer`).
- **Database**: Single Excel workbook (`shared/CallerID-Database.xlsx`) managed via `exceljs`.
  - Concurrency & journal protection (`.pending-changes.jsonl` with cross-process file locks).
  - Works with real Excel desktop or browser live view without file locking conflicts.
  - **Cloud/Vercel Ready**: Automatically boots from `/tmp/shared/` in read-only serverless environments.
- **Frontend**: Pure Vanilla HTML5, modern CSS3 (Mobileum design system, sleek dark glassmorphism, responsive cards, micro-animations), and modular vanilla JavaScript.
- **Localization**: Philippine Peso (`₱15,000.00` with 12% VAT) and E.164 +63 phone number normalizer supporting Mobile, Metro Manila Landlines, Provincial, and Toll-Free numbers.

---

## 🚀 Running Locally

### 1. Install Dependencies
```bash
npm install
```

### 2. Start All Portals (Multi-Port Mode)
```bash
npm start
```
Starts all three portals concurrently:
- Operator Console: [http://localhost:4001](http://localhost:4001)
- Enterprise Portal: [http://localhost:4002](http://localhost:4002)
- Customer Phone: [http://localhost:4003](http://localhost:4003)

### 3. Or Start in Unified Single-Port Mode
```bash
npm run unified
```
Runs all portals on a single unified port: [http://localhost:3000](http://localhost:3000).

---

## 📦 How to Push to GitHub

Follow these steps to push your project to a new GitHub repository:

### Step 1: Initialize Git & Commit
```bash
# Initialize git if not already done
git init

# Stage all files
git add .

# Create the initial commit
git commit -m "Initial commit: Mobileum Caller ID Platform ready for GitHub and Vercel"
```

### Step 2: Create a Repository on GitHub
1. Go to [GitHub.com/new](https://github.com/new).
2. Name your repository (e.g. `mobileum-caller-id-platform`).
3. Set the repository to **Public** or **Private**.
4. Leave "Initialize this repository with a README" **unchecked** (we already created one).
5. Click **Create repository**.

### Step 3: Link & Push
Run the following commands in your terminal (replace `<USERNAME>` and `<REPO_NAME>` with your GitHub details):

```bash
# Rename branch to main
git branch -M main

# Add your GitHub remote repository
git remote add origin https://github.com/<USERNAME>/<REPO_NAME>.git

# Push to GitHub
git push -u origin main
```

---

## ☁️ Deploying to Vercel

The project is pre-configured with `vercel.json` and `api/index.js` for zero-configuration deployment on Vercel.

### Method 1: Import via Vercel Dashboard (Recommended)
1. Go to [vercel.com](https://vercel.com) and log in with your GitHub account.
2. Click **"Add New..."** &rarr; **"Project"**.
3. Select your newly pushed GitHub repository (`mobileum-caller-id-platform`) and click **"Import"**.
4. In the Project Configuration:
   - **Framework Preset**: Other (automatically detected)
   - **Root Directory**: `./` (default)
   - **Build and Output Settings**: Leave as default (configured in `vercel.json`)
5. Click **"Deploy"**.
6. Vercel will build and deploy the application within ~30 seconds, providing you with a live URL (e.g. `https://mobileum-caller-id-platform.vercel.app`).

### Method 2: Deploy via Vercel CLI
If you have the Vercel CLI installed:
```bash
npm i -g vercel
vercel
```
Follow the interactive prompts and select default settings. For production:
```bash
vercel --prod
```

### How Vercel Deployment Works Under the Hood
- **Static Assets**: All portal pages (`/`, `/operator`, `/enterprise`, `/customer`, `/shared/*`) are served directly from Vercel's global CDN Edge for instant page loads.
- **Serverless API**: All API calls (`/api/*`) are handled by `api/index.js` running Express.
- **Serverless Storage**: On Vercel, the read-only file system restriction is handled automatically by copying the bundled `CallerID-Database.xlsx` template and sample logos to `/tmp/shared/` on cold start, flushing transactions synchronously without `EROFS` errors.

---

## 🎬 Live Demo Walkthrough Script

1. **Step 1 — Operator Console ([http://localhost:4001](http://localhost:4001) or `/operator`)**:
   - Go to **Published Packages** &rarr; click **+ New Package**.
   - Create a plan (e.g. `Corporate Gold`, Capacity: `10` numbers, Price: `₱25,000.00`).
   - Click **Publish Package**. The plan is now live in the Excel database.

2. **Step 2 — Enterprise Portal ([http://localhost:4002](http://localhost:4002) or `/enterprise`)**:
   - Go to the **Register** tab.
   - Use the **Multi-Step Wizard**:
     - **Step 1: Choose Package**: Select your newly published package. Click **Next: Brand Identity &rarr;**.
     - **Step 2: Brand Identity**: Enter official business name (e.g. `Bank of the Philippine Islands`), optional tagline, and upload an official logo. Click **Next: Departments & Numbers &rarr;**.
     - **Step 3: Departments & Numbers**: Click quick preset chips (e.g. `Customer Service`, `Fraud & Security`) or click **Fill Sample Numbers**. Observe live number validation and capacity check.
     - Click **Submit Registration for Approval &rarr;**.

3. **Step 3 — Customer Phone Simulator ([http://localhost:4003](http://localhost:4003) or `/customer`)**:
   - In the smartphone simulator, dial the submitted department number (e.g. `+63 917 123 4567`).
   - Because the registration is still pending approval, the screen displays **"Unverified Preview / Pending Verification"**.

4. **Step 4 — Operator Approval**:
   - Return to the **Operator Console**.
   - In **Pending Approvals**, click the new application.
   - Click **Approve Application**.

5. **Step 5 — Real-time Verification on Phone**:
   - Return to the **Customer Phone Simulator** and dial the number again.
   - The smartphone immediately shows the verified green badge (`✓ Verified business`), official bank name, department, tagline, and corporate logo!

6. **Step 6 — Live Excel Database**:
   - Open **Live Excel Sheet** ([http://localhost:4001/sheet.html](http://localhost:4001/sheet.html) or `/sheet.html`).
   - Notice the new rows automatically written into the `Packages`, `Requests`, `Numbers`, and `Files` sheets!

---

## 📄 License
Mobileum Caller ID Platform Demonstration. Designed for Telco Compliance in the Philippines.
