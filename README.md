# 📤 Export Distribution List Members with Microsoft Graph API

This PowerShell script exports all members of **Distribution Lists (DLs)** in a Microsoft 365 tenant using **Microsoft Graph API** via **App-only authentication**. Members are saved in Excel files, and **guest users are separated** into a dedicated folder for auditing or review.

---

## ✅ Purpose

This script is designed for IT administrators who want to:

- Export all **mail-enabled, non-security groups** (classic DLs).
- Retrieve member lists (including nested members).
- Separate **guest users** into a distinct export for compliance checks.
- Use **App-only tokens** for automated, credential-free execution.

---

## 📦 Features

- Uses **App-only authentication** via `MSAL.PS` (no user login required).
- Filters to only **Distribution Lists** (not Microsoft 365 or security groups).
- Follows `@odata.nextLink` pagination for large groups.
- Exports data to `.xlsx` using `ImportExcel`.
- **Guest members are saved in a separate folder**.

---
## 🛠️ Requirements
```powershell
Install-Module MSAL.PS -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
