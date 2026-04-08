# RRM Email Automation Software

Professional Electron-based desktop application for automated bulk email dispatch with intelligent file attachment filtering.

## 🚀 Overview

This tool streamlines the process of sending high volumes of emails based on local folder structures. It parses folder names to extract metadata, generates customized subjects, and applies sophisticated recursive filtering rules to gather necessary attachments automatically.

## ✨ Key Features

### **1. Intelligent Folder Parsing**

- **Bulk Import**: Select a root directory and the app will instantly parse all subfolders.
- **Regex Extraction**: Automatically extracts **Location, Cattle Count, Date,** and **Customer Name** from standardized folder names (e.g., `BOI MARTOLI (10 CATTLE) Dt. 18-02-2026 - RABARI HETALBEN...`).
- **Flexible Subject Generation**: Handles complex cases like "Claim" folders by searching for specific inner folder names to use as subjects.

### **2. Advanced Recursive Attachment Engine**

The app features a deep-traversal scanner with type-specific exclusion rules:

- **🟦 CLAIM**: Scans the entire folder tree but **excludes** any folder named `"Extra"` (case-insensitive).
- **🟡 RETAGGING**: Scans recursively but **excludes** any file whose name contains `"DETAIL"`.
- **🟢 TAGGING**: Scans recursively, excludes `"HC"` folders and `"DETAIL"` files, and enforces a strict whitelist of extensions (.pdf, .jpg, .png, .xls).

### **3. Professional SMTP Core**

- **Secure Configuration**: Built-in Settings panel (Gear ⚙️ icon) for SMTP host, port, and credentials.
- **Local Persistence**: Settings are saved in `rrm-config.json` in the user data directory for recurring use.
- **Batch Dispatch**: Orchestrates large mailing lists with a **1-second delay** between emails to prevent server throttling.
- **Status Tracking**: Visual indicators for `Pending`, `Sending`, `Sent ✅`, and `Error ❌`.

### **4. Premium Responsive UI**

- **Unified Grid Layout**: Every email preview is pixel-perfect and stable, using a custom CSS Grid architecture to prevent field overlapping.
- **Single-Item Carousel**: Easily navigate through your mailing list with "Previous/Next" controls.
- **Glassmorphism Design**: High-end aesthetic with vibrant gradients, custom scrollbars, and animated form inputs.

## 🛠️ Technology Stack

- **Frontend**: React.js with Vanilla CSS.
- **Runtime**: Electron (Node.js backend + Chromium frontend).
- **Email Engine**: Nodemailer (SMTP transport).
- **Build System**: Vite with Electron-Vite.

## 📋 usage

1. **Browse**: Click "Browse" and select your root folder containing the cattle records.
2. **Setup SMTP**: Click the **Gear icon** in the top right and enter your email server details.
3. **Review**: Use the "Previous/Next" buttons to check the parsed data and attachments.
4. **Send**: Click **"Send All Emails"** to start the automated dispatch process.

---

_Developed for advanced email automation and cattle insurance record keeping workflow._
