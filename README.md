# Multi-Brand Inventory & Logistics Portal

Production-ready web application for multi-channel inventory management, inward batch tracking, platform dispatch logging (Blinkit, Zepto, Amazon, Instamart, GT, MT), courier freight & weight audits, and consumer offtake analytics.

---

## ?? Repository Structure

`
+-- app.py                     # Main Flask web application & REST API endpoints
+-- inventory.db               # SQLite database (pre-seeded with schema, data, brands, roles)
+-- requirements.txt           # Python dependencies (Flask, openpyxl, gunicorn)
+-- Procfile                   # Cloud process file (web: gunicorn app:app)
+-- templates/
¦   +-- index.html             # Single-page responsive web dashboard (Tailwind CSS, FontAwesome)
+-- uploads/
¦   +-- dispatch_docs/         # Stored PO copies, invoices, and delivery challans
+-- start_inventory_portal.bat # Windows 1-click launcher
+-- README.md                  # Developer & deployment documentation
`

---

## ?? Quick Start (Local Development)

### 1. Prerequisites
- Python 3.10 or higher
- pip package manager

### 2. Install Dependencies
`ash
pip install -r requirements.txt
`

### 3. Run Application
`ash
python app.py
`
Open **http://127.0.0.1:8765** in your browser.

---

## ?? Production Deployment

The project is pre-configured with gunicorn, SQLite WAL mode, and dynamic $PORT binding for standard hosting platforms.

### Option A: Cloud PaaS (Render / Railway / Heroku) — Easiest
1. Push this repository to **GitHub** or **GitLab**.
2. Create a new **Web Service** on [Render.com](https://render.com) or [Railway.app](https://railway.app).
3. Connect your repository.
4. Set:
   - **Build Command**: pip install -r requirements.txt
   - **Start Command**: gunicorn app:app (or leave default, picked up from Procfile)
5. Click **Deploy**. The platform automatically assigns a free public URL with SSL (https://...).

> **Note on Persistent Storage**: If hosting on free cloud tiers with ephemeral disks, attach a persistent disk for inventory.db and the uploads/ folder so changes persist across restarts.

---

### Option B: Linux VPS (Ubuntu / Debian / AWS EC2 / DigitalOcean)

#### 1. Setup Virtual Environment & Dependencies
`ash
sudo apt update && sudo apt install -y python3-venv python3-pip
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
`

#### 2. Systemd Service (/etc/systemd/system/inventory.service)
`ini
[Unit]
Description=Inventory Management Portal
After=network.target

[Service]
User=www-data
WorkingDirectory=/var/www/inventory
Environment="PATH=/var/www/inventory/venv/bin"
ExecStart=/var/www/inventory/venv/bin/gunicorn --workers 3 --bind 127.0.0.1:8000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
`

#### 3. Enable & Start Service
`ash
sudo systemctl daemon-reload
sudo systemctl enable inventory
sudo systemctl start inventory
`

#### 4. Nginx Reverse Proxy Config (/etc/nginx/sites-available/inventory)
`
ginx
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
`
Enable site and restart Nginx:
`ash
sudo ln -s /etc/nginx/sites-available/inventory /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl restart nginx
`

---

## ?? Security Configuration
- **Max Upload Limit**: 16 MB enforced via Flask configuration.
- **Security Headers**: X-Frame-Options: SAMEORIGIN, X-Content-Type-Options: nosniff.
- **Database Concurrency**: SQLite configured with PRAGMA journal_mode=WAL for concurrent multi-worker reading and writing.
