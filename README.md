# Multi-Brand Inventory & Logistics Portal

Production-ready enterprise web application for multi-channel inventory management, inward batch tracking, platform dispatch logging (Blinkit, Zepto, Amazon, Instamart, GT, MT), courier freight & weight audits, and consumer offtake analytics.

---

## 📁 Repository Structure

```
├── app.py                     # Main Flask web application & REST API endpoints
├── inventory.db               # SQLite database (pre-seeded with schema, users, brands, roles)
├── requirements.txt           # Python dependencies (Flask, openpyxl, gunicorn)
├── Procfile                   # Cloud process file (web: gunicorn app:app)
├── templates/
│   ├── index.html             # Operations dashboard (Hub launcher + modules)
│   └── login.html             # Enterprise login portal with quick credentials drawer
├── uploads/
│   └── dispatch_docs/         # Stored PO copies, invoices, and delivery challans
├── test_auth.py               # Automated unit tests for authentication & security
├── start_inventory_portal.bat # Windows 1-click launcher
└── README.md                  # Developer & deployment documentation
```

---

## 🔐 Authentication & Access Control

The portal is secured with **server-side session authentication** using HTTP-Only cookies and cryptographic password hashing (scrypt with salt via werkzeug.security). Client-controlled headers are strictly blocked.

### Default Team Credentials
All initial accounts are pre-seeded with default password: **Admin@123**

| Role | Email | Permissions |
|---|---|---|
| **Administrator** | admin@company.com | Full system access, platform assignments, audit log, all sheets |
| **Inventory Lead** | inventory@company.com | Master stock baselines, batch inwarding |
| **Joint Lead** | joint_gt@company.com | Dual role: Master stock baselines + General Trade (GT) dispatches |
| **All-Dispatch Lead** | dispatch@company.com | All platform dispatches (Online, GT, MT) |
| **Online Dispatch Lead** | dispatch_online@company.com | 10 Online platforms (Blinkit, Zepto, Instamart, etc.) |
| **GT Dispatch Lead** | dispatch_gt@company.com | General Trade (GT) B2B logistics |
| **MT Dispatch Lead** | dispatch_mt@company.com | Modern Trade (MT) retail chain logistics |
| **Platform Leads** | person1@company.com, person2@company.com, person3@company.com | Courier tracking, freight costs & sales upload for assigned channels |
| **Guest Viewer** | viewer@company.com | Read-only mode across all dashboards |

> **Password Management**: Users can update their password anytime by clicking the key icon next to their profile pill in the header. Administrators can also assign new team members from the **Platform Assignments** modal, which automatically provisions their account with Admin@123.

---

## 🚀 Quick Start (Local Development)

### 1. Prerequisites
- Python 3.10 or higher
- pip package manager

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run Automated Security & Auth Tests
```bash
python test_auth.py
```

### 4. Run Application
```bash
python app.py
```
Open **http://127.0.0.1:8765** in your browser and sign in with any of the accounts above.

---

## 🌐 Production Deployment

The project is pre-configured with gunicorn, SQLite WAL mode, and dynamic $PORT binding for standard hosting platforms.

### Option A: Cloud PaaS (Render / Railway / Heroku) - Easiest
1. Push this repository to **GitHub** or **GitLab**.
2. Create a new **Web Service** on Render.com or Railway.app.
3. Connect your repository.
4. Set:
   - **Build Command**: pip install -r requirements.txt
   - **Start Command**: gunicorn app:app (or leave default, picked up from Procfile)
5. Configure Environment Variables:
   - `SECRET_KEY`: Set to a strong random secret key for session cookie encryption.
6. Click **Deploy**. The platform automatically assigns a free public URL with SSL (https://...).

> **Note on Persistent Storage**: If hosting on cloud tiers with ephemeral disks, attach a persistent disk for inventory.db and the uploads/ folder so changes persist across restarts.

---

### Option B: Linux VPS (Ubuntu / Debian / AWS EC2 / DigitalOcean)

#### 1. Setup Virtual Environment & Dependencies
```bash
sudo apt update && sudo apt install -y python3-venv python3-pip
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

#### 2. Systemd Service (/etc/systemd/system/inventory.service)
```ini
[Unit]
Description=Inventory Management Portal
After=network.target

[Service]
User=www-data
WorkingDirectory=/var/www/inventory
Environment="PATH=/var/www/inventory/venv/bin"
Environment="SECRET_KEY=your-production-secret-key-here"
ExecStart=/var/www/inventory/venv/bin/gunicorn --workers 3 --bind 127.0.0.1:8000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

#### 3. Enable & Start Service
```bash
sudo systemctl daemon-reload
sudo systemctl enable inventory
sudo systemctl start inventory
```

#### 4. Nginx Reverse Proxy Config (/etc/nginx/sites-available/inventory)
```nginx
server {
    listen 80;
    server_name your-domain.com;
    client_max_body_size 16M;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```
Enable site and restart Nginx:
```bash
sudo ln -s /etc/nginx/sites-available/inventory /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl restart nginx
```

---

## 🛡️ Security Hardening Implemented
- **Server-Side Session Cookies**: HTTP-Only, SameSite=Lax, encrypted with server SECRET_KEY.
- **Defense Against Spoofing**: Client-side headers (X-User-Email) are completely ignored; identity is strictly tied to validated server sessions.
- **Scrypt Password Hashing**: Passwords stored as salted scrypt hashes via werkzeug.security.
- **Brute Force & Input Protection**: Strict 16MB upload limit, sanitized file uploads, centralized @app.before_request login barrier.
- **HTTP Security Headers**: X-Frame-Options: SAMEORIGIN, X-Content-Type-Options: nosniff, X-XSS-Protection: 1; mode=block.
- **Database Concurrency**: SQLite configured in WAL (Write-Ahead Logging) mode with 30-second busy timeout.
