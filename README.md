# Deeksharambh 2026 — Deployment Guide
**Flask + MongoDB Atlas · Office of Academics · JAIN University**

---

## Quick Start (Local)

```bash
pip install -r requirements.txt
cp .env.example .env          # fill in MONGO_URI, SECRET_KEY, etc.
python app.py
```

Visit `http://localhost:5000` for the student form.
Visit `http://localhost:5000/admin` for the dashboard.

---

## Production Deployment (Hostinger VPS / Ubuntu 24.04)

### 1. Copy files to server
```bash
scp -r deeksharambh/ user@31.97.186.191:/var/www/deeksharambh
```

### 2. On the server
```bash
cd /var/www/deeksharambh
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
nano .env   # Fill in MONGO_URI, SECRET_KEY, ADMIN_PASSWORD
```

### 3. Add your JAIN logo
Replace the placeholder with your actual logo:
```bash
# Place your JAIN logo at:
/var/www/deeksharambh/static/img/logo.png
# Recommended: PNG, white/transparent background, ~200×60px
```

### 4. Start with systemd
```bash
sudo cp deeksharambh.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable deeksharambh
sudo systemctl start deeksharambh
sudo systemctl status deeksharambh
```

### 5. Nginx reverse proxy
```bash
sudo cp nginx.conf.example /etc/nginx/sites-available/deeksharambh
# Edit your domain name in the file, then:
sudo ln -s /etc/nginx/sites-available/deeksharambh /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

### 6. SSL (Certbot)
```bash
sudo certbot --nginx -d your-domain.com
```

---

## Project Structure

```
deeksharambh/
├── app.py                  ← Flask app (routes, MongoDB, Excel export)
├── requirements.txt
├── gunicorn.conf.py        ← Production Gunicorn config
├── deeksharambh.service    ← Systemd unit
├── nginx.conf.example      ← Nginx reverse-proxy config
├── .env.example            ← Copy to .env and fill in
├── templates/
│   ├── form.html           ← Student survey (landing page + GSAP)
│   ├── admin.html          ← Admin dashboard (live charts)
│   └── admin_login.html    ← Admin login
└── static/
    └── img/
        └── logo.png        ← ← ← REPLACE WITH YOUR ACTUAL JAIN LOGO
```

---

## API Endpoints

| Method | URL | Auth | Description |
|--------|-----|------|-------------|
| GET | `/` | Public | Student survey (landing page) |
| POST | `/api/submit` | Public | Submit response (rate-limited 5/min/IP) |
| GET | `/admin/login` | — | Admin login page |
| GET | `/admin/dashboard` | Admin | Live dashboard |
| GET | `/admin/api/stats` | Admin | JSON response data |
| GET | `/admin/api/schools` | Admin | Unique schools list |
| GET | `/admin/api/export` | Admin | Download Excel (.xlsx) |
| POST | `/admin/api/delete_selected` | Admin | Delete responses by ID |
| POST | `/admin/api/delete_all` | Admin | Delete all responses |
| GET | `/health` | Public | DB health check |

---

## Form Design Notes

- **Landing page** with GSAP letter-by-letter animation, feature cards, and timeline
- **10-step survey** with slide-in/out animations and staggered Q-card entrances
- **Star rating** for Q2 (vibe score 1–10), emoji cards for Q5/Q10/Q14/Q15/Q20/Q22/Q25/Q35
- ~~**PDF download** on thank-you screen~~ (removed)
- **localStorage** used for "already submitted" detection (no server-side session on form)
- **Q5** stores the emoji label text (e.g. "Smooth"); `app.py` maps it to numeric and stores both `q5` (numeric) and `q5_label` (text)

## Admin Dashboard Notes

- **12 tabs**: Overview, Student List, Vibe Check, Transition, Footsteps, Orientation, Bridge Course, NEP, Gen Z, Belonging, NPS & Score, By School
- **Emoji-to-numeric** mapping in JS for Q5, Q14, Q15, Q20, Q22, Q25, Q35 distribution charts
- **Tone badge** in student list (Positive / Critical / Mixed) based on Q1 word selection
- **Q5a / Q5b** (spark vs missed sessions) shown as separate bar charts in Vibe Check tab
- **Auto-refresh** every 90 seconds

## MongoDB Atlas Setup

1. Create free M0 cluster at cloud.mongodb.com
2. Add IP `0.0.0.0/0` to Network Access
3. Create DB user with read/write permissions
4. Copy connection string to `.env` as `MONGO_URI`

Collection: `deeksharambh2026.responses`

## Logs

```bash
sudo journalctl -u deeksharambh -f
sudo tail -f /var/log/nginx/access.log
```
