# 1997-I-Ulamudheen-Higher Secondary Madrassa Election 2026

A modern, entirely free voting web application designed for a 700+ student body without requiring standalone server infrastructure. Uses plain HTML/JS and relies on Google Sheets API as an exceptionally reliable backend data-store!

## Step 1. Google Sheets Backend Setup
1. Create a new blank Google Sheet. You can name it "Madrassa Election DB".
2. Go to **Extensions > Apps Script**.
3. Clear out whatever template code is there and copy everything from the `apps-script.gs` file in this folder into the code editor entirely.
4. Click the **Save** (floppy disk) icon.
5. Click the **Deploy** button on the top right, then select **New deployment**.
6. Click the gear icon next to "Select type" and choose **Web app**.
   - **Description:** "Election DB"
   - **Execute as:** "Me"
   - **Who has access:** "Anyone"
7. Click **Deploy**. Google may ask you to authorize access. Proceed through the Advanced security warning carefully to allow your script access.
8. Under "Web app URL", click "Copy".
9. Open `candidates.js` in this folder, locate `const BACKEND_URL = "";` and paste this copied URL strictly **inside the quotes**.

## Step 2. How to Change Candidate Names and Icons
All customizations are done without any complex coding. Simply edit the `candidates.js` file located at the main folder level!
Open it, and find the `POSITIONS` array block.

To edit a candidate:
```javascript
// BEFORE
{ name: "Izzudheen EK", icon: "💻", booth: "both" }
// AFTER (Example with custom icon file)
{ name: "Abdullah Ali", icon: "icons/my-new-icon.png", booth: "boys" }
```
- Put custom image files into the `/icons/` folder and name them exactly as mapped. Emojis work fine out-of-the box.

## Hosting for Free
Simply upload all these files (the HTML files, JS, CSS, and Icons) to a web host such as GitHub Pages, Netlify Drop, or standard file-hosting like Hostinger. Since it relies on Google APIs natively, the App functions entirely off static client-side rendering.
