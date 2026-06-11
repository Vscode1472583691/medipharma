# MediPharma Deployment Guide - Render.com

## Overview
This guide will help you deploy the MediPharma Flask application to **Render.com** with a MySQL database.

---

## Prerequisites

1. **GitHub Account** - Push your code to GitHub first
2. **Render.com Account** - Sign up at https://render.com (free tier available)
3. **Git installed** - For version control

---

## Step 1: Prepare Your Code for Deployment

### 1.1 Create a .env file for production variables

Create a `.env` file in your project root (for local testing):

```bash
FLASK_ENV=production
SECRET_KEY=your-very-secure-random-key-here
```

**Important**: Never commit `.env` to GitHub. It's in `.gitignore`.

### 1.2 Verify production files exist:
- ✅ `requirements.txt` - Updated with `gunicorn` and `PyMySQL`
- ✅ `render.yaml` - Deployment configuration
- ✅ `wsgi.py` - WSGI entry point for production

---

## Step 2: Push Code to GitHub

```bash
# Initialize git (if not already done)
git init

# Add all files
git add .

# Commit
git commit -m "Prepare for Render deployment"

# Add remote (replace with your GitHub repo URL)
git remote add origin https://github.com/your-username/medipharma.git

# Push to GitHub
git branch -M main
git push -u origin main
```

**Important**: Make sure you have a `.gitignore` file:

```
venv/
*.db
.env
__pycache__/
*.pyc
dist/
build/
*.spec
.pytest_cache/
```

---

## Step 3: Deploy on Render.com

### 3.1 Create a Render Account
1. Go to https://render.com
2. Sign up with GitHub (recommended)
3. Connect your GitHub account

### 3.2 Deploy Your App
1. Click **"New"** → **"Web Service"**
2. Select your GitHub repository
3. Fill in the form:
   - **Name**: `medipharma` (or your preferred name)
   - **Runtime**: `Python 3.11`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
   - **Plan**: Free (or paid if you want better performance)

### 3.3 Set Environment Variables
In the Render dashboard, go to **Environment** and add:

```
FLASK_ENV=production
SECRET_KEY=<generate-a-strong-random-key>
DATABASE_URL=<will-be-set-by-MySQL-service>
```

To generate a strong SECRET_KEY, run in Python:
```python
import secrets
print(secrets.token_hex(32))
```

### 3.4 Create MySQL Database
1. In Render dashboard, click **"New"** → **"MySQL Database"**
2. Fill in:
   - **Name**: `medipharma-db`
   - **Database Name**: `medipharma`
   - **User**: `medipharma_user`
   - **Region**: Same as your web service
   - **Plan**: Free or paid

3. After creation, copy the **Internal Database URL** from the MySQL service details
4. Add it to your web service as environment variable:
   ```
   DATABASE_URL=<paste-the-internal-database-url>
   ```

---

## Step 4: Database Migration (First Time Only)

After your app is deployed, you need to initialize the database tables:

### Option A: Using Render Shell

1. In Render dashboard, go to your web service
2. Click **"Shell"** tab
3. Run these commands:

```bash
python
>>> from app import app, db
>>> with app.app_context():
>>>     db.create_all()
>>>     print("Database initialized!")
>>> exit()
```

### Option B: Create a Migration Script

Create `init_db.py`:

```python
from app import app, db

with app.app_context():
    db.create_all()
    print("Database tables created successfully!")
```

Then add to `render.yaml`:
```yaml
preDeployCommand: "python init_db.py"
```

---

## Step 5: Verify Your Deployment

1. Render will provide you with a URL like: `https://medipharma.onrender.com`
2. Open the URL in your browser
3. Login with default credentials:
   - **Username**: `admin`
   - **Password**: `admin1472`

---

## Step 6: Update Default Credentials (IMPORTANT!)

⚠️ **Security**: Change the default admin password immediately!

1. Log in with default credentials
2. Go to **Dashboard** → **Change Password**
3. Set a strong new password

---

## Troubleshooting

### App not starting?

Check logs in Render dashboard:
```bash
# In Render Shell or Logs tab
tail -f logs.log
```

### Database connection error?

Verify environment variables:
- Check `DATABASE_URL` is correctly set
- Ensure MySQL service is running
- Wait 5-10 minutes for services to fully initialize

### Static files not loading?

Flask automatically serves static files. If issues persist:

```python
# In app.py - ensure this is correct
app = Flask(__name__, 
    template_folder=template_folder, 
    static_folder=static_folder)
```

### Permission denied errors?

Make sure you have a `PYTHONUNBUFFERED=1` environment variable set in Render.

---

## Domain Setup (Optional)

To use a custom domain:

1. In Render dashboard, go to your service
2. Click **Settings** → **Custom Domain**
3. Add your domain name
4. Update DNS records with the CNAME provided

---

## Backup and Recovery

### Backup Database

From Render Shell:
```bash
# This depends on your MySQL setup
# Check Render documentation for backup procedures
```

### Restore from Local Database

Export data from local SQLite, then import to MySQL using a migration script.

---

## Costs (as of 2024)

- **Web Service**: Free tier (shared CPU, 0.5GB RAM)
- **MySQL Database**: Free tier (1GB storage)
- **Paid upgrades**: Available if needed for higher traffic

---

## Support & Next Steps

- 📚 Render Docs: https://render.com/docs
- 🐛 Flask Docs: https://flask.palletsprojects.com/
- 💬 Get help: Check Render community forums

---

## Maintenance

### Keep Dependencies Updated

Periodically update Python packages:

```bash
pip list --outdated
# Then update requirements.txt and push to GitHub
```

### Monitor Performance

In Render dashboard:
- Check **Metrics** for CPU/Memory usage
- Review **Logs** for errors
- Monitor database size

### Regular Backups

Set up automated backups for your MySQL database through Render.

---

Good luck with your deployment! 🚀
