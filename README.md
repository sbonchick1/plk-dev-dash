# Popeyes Restaurant Openings Dashboard

A live dashboard displaying Popeyes restaurant opening data from Smartsheet with interactive filters and analytics.

## 📁 Repository Structure
```
popeyes-dashboard/
├── api/
│   └── sheet.js          # Serverless API function for Smartsheet data
├── public/
│   └── index.html        # Dashboard frontend with charts and filters
├── package.json          # Node.js dependencies
├── vercel.json           # Vercel deployment configuration
└── README.md             # This file
```

## 🚀 Deployment

This dashboard is deployed on Vercel and connects to Smartsheet API.

### Environment Variables Required:
- `SMARTSHEET_TOKEN` - Your Smartsheet API access token

### Deploy to Vercel:
1. Connect this repository to Vercel
2. Add the `SMARTSHEET_TOKEN` environment variable
3. Deploy!

## 🔧 Technologies Used

- **Frontend**: HTML, CSS, JavaScript, Chart.js
- **Backend**: Node.js, Vercel Serverless Functions
- **Data Source**: Smartsheet API

## 📊 Features

- Live data from Smartsheet
- Interactive filtering by Year, Division, Architecture Type, Urbanicity, and FSS Grade
- Multiple dashboard views: Overview, Division, Architecture, Urbanicity, FSS
- Real-time KPI metrics
- Visual charts and graphs

## 🔄 Auto-Deploy

Pushing to the main branch automatically triggers a new deployment on Vercel.
