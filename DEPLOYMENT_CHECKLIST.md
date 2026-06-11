# MediPharma Deployment Checklist

## Pre-Deployment ✓

- [ ] **Git Setup**
  - [ ] Install Git: https://git-scm.com
  - [ ] Create GitHub account: https://github.com
  - [ ] Create new repository named `medipharma`

- [ ] **Code Preparation**
  - [ ] Verify `requirements.txt` has all dependencies
  - [ ] Verify `.gitignore` exists
  - [ ] Verify `.env.example` exists
  - [ ] Create `.env` for local testing (not for GitHub)
  - [ ] Verify `render.yaml` exists
  - [ ] Verify `wsgi.py` exists

## GitHub Push ✓

```bash
# From your project directory
git init
git add .
git commit -m "Initial commit - Ready for deployment"
git remote add origin https://github.com/YOUR-USERNAME/medipharma.git
git branch -M main
git push -u origin main
```

**Copy your GitHub repo URL** - you'll need it for Render

## Render Deployment ✓

1. **Create Account** - Sign up at https://render.com

2. **Deploy Web Service**
   - [ ] Click "New" → "Web Service"
   - [ ] Connect your GitHub repo
   - [ ] Name: `medipharma`
   - [ ] Build: `pip install -r requirements.txt`
   - [ ] Start: `gunicorn app:app --bind 0.0.0.0:$PORT`
   - [ ] Plan: Free
   - [ ] Click "Create Web Service"

3. **Add Environment Variables** (in Render Dashboard)
   ```
   FLASK_ENV=production
   SECRET_KEY=<run: python -c "import secrets; print(secrets.token_hex(32))">
   ```

4. **Create MySQL Database** (in Render Dashboard)
   - [ ] Click "New" → "MySQL"
   - [ ] Name: `medipharma-db`
   - [ ] Database Name: `medipharma`
   - [ ] User: `medipharma_user`
   - [ ] Plan: Free
   - [ ] Create Database

5. **Connect Database to Web Service**
   - [ ] Copy MySQL "Internal Database URL"
   - [ ] Go to Web Service → Environment
   - [ ] Add: `DATABASE_URL=<paste-url>`
   - [ ] Click "Save"

## Post-Deployment ✓

1. **Initialize Database**
   - [ ] Wait 2-3 minutes for services to start
   - [ ] In Web Service, click "Shell"
   - [ ] Run:
     ```bash
     python
     >>> from app import app, db
     >>> with app.app_context():
     >>>     db.create_all()
     >>> exit()
     ```

2. **Verify Application**
   - [ ] Copy your Render URL (e.g., https://medipharma.onrender.com)
   - [ ] Open in browser
   - [ ] Login with: admin / admin1472
   - [ ] Test basic features

3. **Security**
   - [ ] Change default admin password immediately
   - [ ] Update `SECRET_KEY` to be truly random
   - [ ] Review database credentials

## Future Updates ✓

When you make code changes:

```bash
git add .
git commit -m "Description of changes"
git push origin main
```

Render automatically re-deploys when you push to GitHub!

## Support URLs

- GitHub: https://github.com
- Render: https://render.com
- MySQL: https://dev.mysql.com
- Flask: https://flask.palletsprojects.com
- Python: https://python.org

## Quick Tips

⚡ **Speed up redeploys**: Use `git commit --amend` to edit last commit
🔄 **Auto-redeploy**: Enabled by default - commits trigger rebuilds
📊 **Monitor**: Check Render dashboard Logs tab for errors
💾 **Backup**: Render provides MySQL backup options

---

**Status**: Ready to Deploy! Follow steps above in order. 🚀
