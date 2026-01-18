# Office to PDF Converter

A free, fast, and secure web application to convert Excel & PowerPoint files to PDF format with a beautiful warm-themed interface.

## Features

- 📊 Convert Excel files (.xlsx, .xls, .xlsm) to PDF
- 📽️ Convert PowerPoint files (.pptx, .ppt) to PDF
- 🔒 Secure - files are not stored on the server
- ⚡ Fast conversion
- 💯 Completely free with no limits
- 📱 Responsive design - works on all devices
- 🎨 Beautiful, luxury UI with smooth animations
- ✨ Warm gradient theme (yellow, orange, red, purple)

## Local Development

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### Installation

1. Clone or navigate to this directory:
```bash
cd h:\pdfconverter
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
```

3. Activate the virtual environment:
```bash
# Windows
.\venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

### Running Locally

```bash
python app.py
```

The application will be available at `http://localhost:5000`

## Deployment Options

### Option 1: Deploy to Render (Recommended - Free Tier Available)

1. Create a free account at [render.com](https://render.com)
2. Click "New +" → "Web Service"
3. Connect your GitHub repository (push this code to GitHub first)
4. Configure:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Environment**: Python 3
5. Click "Create Web Service"

### Option 2: Deploy to Railway

1. Create account at [railway.app](https://railway.app)
2. Click "New Project" → "Deploy from GitHub repo"
3. Select your repository
4. Railway will auto-detect Python and deploy automatically

### Option 3: Deploy to Vercel (Serverless)

1. Install Vercel CLI: `npm i -g vercel`
2. Run `vercel` in the project directory
3. Follow the prompts

### Option 4: Deploy to PythonAnywhere (Free Tier)

1. Create account at [pythonanywhere.com](https://www.pythonanywhere.com)
2. Upload files via Files tab
3. Create a new web app with Flask
4. Configure WSGI file to point to your app

### Option 5: Docker Deployment

Use the included Dockerfile to deploy to any cloud platform that supports Docker.

## Project Structure

```
pdfconverter/
├── app.py                 # Flask backend application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main HTML page
├── static/
│   ├── style.css         # Stylesheet
│   └── script.js         # Frontend JavaScript
├── README.md             # This file
├── Procfile              # For Heroku/Render deployment
├── vercel.json           # For Vercel deployment
└── Dockerfile            # For Docker deployment
```

## Technologies Used

- **Backend**: Flask (Python)
- **PDF Generation**: ReportLab
- **Excel Processing**: openpyxl
- **Frontend**: Vanilla HTML/CSS/JavaScript
- **Styling**: Custom CSS with gradient theme

## Security Notes

- Files are processed in memory and not saved to disk
- Maximum file size: 16MB
- Only Excel file formats are accepted
- CORS enabled for API access

## License

Free to use and modify for personal and commercial projects.

## Support

For issues or questions, please create an issue in the repository.
