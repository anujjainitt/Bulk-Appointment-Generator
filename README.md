# Bulk Appointment Letters

A fullstack application for generating and downloading bulk appointment letters from Excel data. Built with React (frontend) and Node.js (backend).

## Project Structure

```
bulk-appointment-letters/
  Dockerfile
  backend/
    server.js
    package.json
    templates/
    uploads/
  frontend/
    package.json
    src/
    public/
```

## Local Development

1. **Install dependencies:**
   - `cd frontend && npm install`
   - `cd ../backend && npm install`
2. **Run frontend:**
   - `cd ../frontend && npm run dev`
3. **Run backend:**
   - `cd ../backend && node server.js`

## Production Build & Deploy (AWS App Runner)

1. **Build frontend:**
   - `cd frontend && npm run build`
2. **Backend serves frontend build automatically.**
3. **Docker build & deploy:**
   - Push to GitHub and connect to AWS App Runner.
   - App Runner will use the Dockerfile to build and run the app.

## Features
- Upload Excel, preview data, and generate appointment letters in bulk.
- Responsive, modern UI (Material UI).
- Download all appointment letters as a zip file.
- Ready for cloud deployment.

--- 