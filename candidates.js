// ============================================
// EDIT YOUR CANDIDATES HERE — NO CODING NEEDED
// ============================================
//
// HOW TO CHANGE THINGS:
// 1. Changing a name: Edit the "name" field. Example: name: "New Name"
// 2. Swapping an icon: Change the "icon" field to an Emoji (e.g. "🍎") OR a file path from the icons folder (e.g. "icons/apple.png")
// 3. Adding a new candidate: Copy an existing block `{ name: ..., icon: ..., booth: ... }` and add it to the correct position list. Don't forget the comma!
// 4. Removing a candidate: Simply delete their block.
// 5. Where to put custom icon files: Place them in the "icons" folder that came with this project. Use them like "icons/filename.png"
//
// BOOTHS:
// Use "boys", "girls", or "both" to control which booth the candidate appears in.

const POSITIONS = [
    {
        title: "Leader",
        candidates: [
            { name: "Izzudheen EK", icon: "icons/Laptop.png", booth: "both" },
            { name: "Abdul Nafih C", icon: "icons/Fan.png", booth: "both" }
        ]
    },
    {
        title: "Deputy Leader",
        candidates: [
            { name: "Musliha Safa MP", icon: "icons/Mug.png", booth: "both" },
            { name: "Minha Nasrin K", icon: "icons/Pen.png", booth: "both" }
        ]
    },
    {
        title: "Fine Arts Convenor",
        candidates: [
            { name: "Muhammed Rafan K", icon: "icons/Table.png", booth: "both" },
            { name: "Muhammed Hadi K", icon: "icons/Bulb.png", booth: "both" }
        ]
    },
    {
        title: "Fine Arts Assistant Convenor",
        candidates: [
            { name: "Fathima Ridha M", icon: "icons/BlackBoard.png", booth: "both" },
            { name: "Afla V", icon: "icons/Bus.png", booth: "both" },
            { name: "Shahima Nooriya M", icon: "icons/Phone.png", booth: "both" }
        ]
    }
];

// ============================================
// GOOGLE SHEETS APP SCRIPT URL LINK
// ============================================
// Follow the README instructions to get this link. Paste it below inside the quotes!
const BACKEND_URL = "https://script.google.com/macros/s/AKfycbzP9B-ycFJslBDPH_ck5igr5bSkboNrEVJXW2K4HH9SLX1RZogfRPmFuzZQD2_rRxL8/exec";