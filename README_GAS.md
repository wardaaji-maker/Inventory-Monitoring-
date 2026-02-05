# Lead Management System (Google Apps Script)

This project contains the source code for the Lead Management System implemented in Google Apps Script.

## Files

- `backend/Code.gs`: The main Google Apps Script backend code.
- `frontend/index.html`: The frontend HTML/JavaScript code.

## Deployment

1. Create a new Google Apps Script project at [script.google.com](https://script.google.com).
2. Copy the content of `backend/Code.gs` into `Code.gs` in the project.
3. Create a new HTML file named `index.html` in the project and copy the content of `frontend/index.html` into it.
4. Deploy as a Web App (Deploy > New Deployment > Web app).
5. Set "Execute as" to "Me" and "Who has access" to "Anyone" (or as appropriate).
6. Copy the "Web App URL" and use it.

## Configuration

Update the `CONFIG` object in `Code.gs` with your Google Sheet ID.
