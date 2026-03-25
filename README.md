# DX-A3 Stock Management System
ระบบจัดการสต็อก Thai Cargo DX-A3

## Architecture
- **Frontend** — Static HTML/CSS/JS hosted on GitHub Pages (this repo)
- **Backend** — Google Apps Script (GAS) as REST API, connected to Google Sheets

```
GitHub Pages (index.html)
        │  fetch() REST calls
        ▼
Google Apps Script (code.gs)  ←→  Google Sheets (Inventory + Movements)
```

## Setup Instructions

### 1. Deploy Google Apps Script as Web App

1. Open [Google Apps Script](https://script.google.com/) and create a new project
2. Copy the contents of `code.gs` into the editor
3. Create a new HTML file named `Index` (File → New → HTML file) — paste `index.html` contents *(only needed for GAS fallback)*
4. Click **Deploy → New deployment**
5. Select type: **Web App**
6. Settings:
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Click **Deploy** and copy the deployment URL

### 2. Configure Frontend

Open `index.html` and replace line:
```javascript
const GAS_API_URL = 'https://script.google.com/macros/s/REPLACE_WITH_YOUR_DEPLOYMENT_ID/exec';
```
with your actual GAS Deployment URL from step 1.

### 3. Add Logo

Save the Thai Cargo logo image as `thai-cargo-logo.png` in the project root directory.

### 4. Enable GitHub Pages

1. Push this repo to GitHub
2. Go to repository **Settings → Pages**
3. Source: **Deploy from a branch → main → / (root)**
4. Save — your app will be live at `https://MajesticERP.github.io/DXA3/`

### 5. Re-deploy GAS after CORS

After GitHub Pages is live, go back to GAS and create a **New Deployment** (or update existing) so the CORS headers are refreshed.

## Files

| File | Description |
|------|-------------|
| `index.html` | Frontend web app (GitHub Pages) |
| `code.gs` | Google Apps Script backend (REST API + Google Sheets) |
| `thai-cargo-logo.png` | Thai Cargo logo *(add manually)* |
| `DX-A3_Store - Inventory.csv` | Inventory data backup |
| `DX-A3_Store - Movements.csv` | Movements data backup |

## Features
- จัดการสต็อกสินค้า (เพิ่ม, แก้ไข, ลบ)
- รับเข้า / เบิกจ่าย พร้อมประวัติ
- แดชบอร์ดภาพรวม + กราฟ
- ส่งออกรายงาน Excel / PDF
- CacheService เพื่อประสิทธิภาพสูง
