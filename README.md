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
---

## 🛠️ Requirements

 PowerShell Modules:
 - [`MSAL.PS`](https://www.powershellgallery.com/packages/MSAL.PS)
- [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel)

  Install required modules:
```powershell
Install-Module MSAL.PS -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```
---

## 🔐 Required Permissions
Your registered app must have the following Microsoft Graph API Application permissions:

- Group.Read.All
- User.Read.All

These require admin consent in the Azure portal.


## Set your Azure AD app credentials in the script:
```powershell
$tenantId     = "<your-tenant-id>"
$clientId     = "<your-client-id>"
$clientSecret = "<your-client-secret>"
```

## 🚀 Usage
Clone or download this script.

Run the script:
```powershell
.\Export-DLMembers.ps1
```
