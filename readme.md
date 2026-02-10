[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# 📊 Google Sheets → Multi-Platform Sync Tool

A powerful Google Apps Script that transforms your Google Sheets into a contact sync hub for major marketing and analytics platforms. Upload contacts with one click, track upload status, and manage duplicate prevention - all without leaving Google Sheets.

> ⚠️ **WORK IN PROGRESS**: This project is currently under active development. Platform integrations are being added progressively. Features and documentation may change. Use at your own discretion.

---

## 📑 Table of Contents
1. [Features](#-features)
2. [Installation & Configuration](#-installation--configuration)
3. [How to Use](#-how-to-use)
4. [Tech Stack](#-tech-stack)
5. [License](#-license)
6. [Author](#-author)

---

## ✨ Features

- **One-Click Uploads** - Sync contacts directly from Google Sheets custom menu
- **Phone Normalization** - Automatic Indonesia phone number formatting (+62 E.164 format)
- **Status Tracking** - Real-time upload status with timestamps
- **Batch Processing** - Handle large datasets efficiently
- **Connection Verification** - Pre-upload API connection testing
- **Upload Statistics** - Visual progress bars and success/failure counts

---

## 🚀 Installation & Configuration

### Step 1: Set Up Google Sheets

1. Create a new Google Sheet or open your existing contact database
2. Ensure your sheet has the following columns:
   - `Name` or `Full Name`
   - `Phone` or `Mobile`
   - `Email`

### Step 2: Add Apps Script

1. In your Google Sheet, go to **Extensions** → **Apps Script**
2. Delete any existing code in the editor
3. Copy the entire script code and paste it into the editor
4. Click **Save** (💾 icon)
5. Name your project (e.g., "Contact Sync Tool")

### Step 3: Configure API Credentials

In the script editor, locate the configuration section and add your API keys:
```javascript
// Example configuration
const CONFIG = {
  PLATFORM_API_KEY: 'your-platform-api-key',
  PLATFORM_ENDPOINT: 'your-platform-endpoint',
  // ... add other platform credentials
};
```

### Step 4: Authorize the Script

1. Click **Run** → Select any function (e.g., `onOpen`)
2. Click **Review Permissions**
3. Choose your Google account
4. Click **Advanced** → **Go to [Project Name] (unsafe)**
5. Click **Allow**

### Step 5: Reload Your Sheet

Close and reopen your Google Sheet. You should see a new menu item called **"Contact Sync"** or similar.

---

## 📖 How to Use

### Upload Contacts to a Platform

1. **Prepare Your Data**
   - Ensure your contact data is properly formatted in the sheet
   - Phone numbers can be in any format (will be auto-normalized for Indonesia)

2. **Select Upload Platform**
   - Click on **Contact Sync** menu in Google Sheets
   - Choose your target platform (e.g., "Upload to Braze")

3. **Confirm Upload**
   - Review the confirmation dialog
   - Click **OK** to proceed

4. **Monitor Progress**
   - Watch the status column update in real-time
   - Check the summary dialog for upload statistics

### Check Connection Status

Before uploading, verify your API connections:

1. Click **Contact Sync** → **Test Connection** → Select platform
2. Review the connection status dialog

### View Recent Contacts

For platforms that support it:

1. Click **Contact Sync** → **View Recent** → Select platform
2. Review the recent contacts list

### Status Indicators

- **Uploaded** - Successfully synced to platform
- **Already in Database** - Duplicate detected, skipped
- **Failed** - Upload error (check logs)
- **Pending** - Queued for upload

---

## 🛠 Tech Stack

- **Google Apps Script** - Automation and scripting
- **UrlFetchApp** - HTTP API requests
- **Batch API Processing** - Efficient bulk operations
- **E.164 Phone Normalization** - International phone formatting
- **Custom Menu UI** - Google Sheets integration

---

## 📝 License

This project is built as a learning reference. You are free to use, modify, and distribute this software for personal learning, testing and demo purposes.

---

## 👤 Author

**Sandi Utomo**  
[![Linkedin](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/sandiutomo/) 
[![Github](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/sandiutomo)