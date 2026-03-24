# Popeyes Dashboard - Vercel Deployment

## 📁 Folder Structure

```
your-repo/
├── api/
│   └── sheet.js          # Serverless API function
├── public/
│   └── index.html        # Dashboard frontend
├── vercel.json           # Vercel configuration
├── package.json          # Dependencies
└── README-VERCEL.md      # This file
```

## 🚀 Deployment Steps

### 1. Add These Files to Your Repository

Copy all the files from this folder into your GitHub repository:
- `api/sheet.js`
- `public/index.html`
- `vercel.json`

Your existing `package.json` and `server.js` can stay as-is.

### 2. Push to GitHub

```bash
git add .
git commit -m "Configure for Vercel deployment"
git push
```

### 3. Deploy to Vercel

1. Go to **https://vercel.com** and sign in with your GitHub account
2. Click **"Add New Project"**
3. Click **"Import Git Repository"**
4. **Select your repository** from the list
5. Configure the project:
   - **Framework Preset**: Other (or leave as auto-detected)
   - **Root Directory**: `./` (leave as default)
   - **Build Command**: Leave empty
   - **Output Directory**: `public`

### 4. Add Environment Variable

In the Vercel project settings:
- Click **"Environment Variables"**
- Add a new variable:
  - **Name**: `SMARTSHEET_TOKEN`
  - **Value**: Your Smartsheet API token
  - **Environments**: Select all (Production, Preview, Development)

### 5. Deploy!

Click **"Deploy"** and Vercel will build your app.

You'll get a URL like: `https://your-app-name.vercel.app`

## 🔄 Future Updates

After the initial deployment, any time you push to GitHub, Vercel will automatically redeploy!

```bash
git add .
git commit -m "Update dashboard"
git push
```

## ⚙️ What Changed?

- **API endpoint**: Changed from `https://popeyes-dashboard.onrender.com/api/sheet` to `/api/sheet`
- **Server architecture**: Express server → Vercel serverless function
- **File structure**: Added `public/` folder and `api/` folder for Vercel's routing

## 🆘 Troubleshooting

**Deploy fails?**
- Make sure `SMARTSHEET_TOKEN` environment variable is set
- Check that all files are in the correct folders
- Verify `package.json` includes `node-fetch` dependency

**Dashboard loads but shows error?**
- Check the Vercel function logs for errors
- Verify your Smartsheet API token is valid
- Test the API endpoint: `https://your-app.vercel.app/api/sheet`

## 📝 Notes

- The `server.js` file is no longer used but can stay in your repo for reference
- Vercel serverless functions have a 10-second execution timeout
- The free tier includes automatic HTTPS and global CDN
