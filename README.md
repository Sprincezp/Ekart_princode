# Pincode Serviceability Checker

A web-based tool to check delivery service availability for Indian pincodes. Supports both road and air service verification with real-time search.

## Features

- 🔍 **Instant Pincode Search** — Type a 6-digit pincode and get results instantly
- 🚚 **Road Services** — Check pickup & delivery availability via road
- ✈️ **Air Services** — Check pickup & delivery availability via air
- 🌗 **Dark/Light Theme** — Toggle between dark and light mode (Ekart2)
- 📱 **Responsive** — Works on mobile, tablet, and desktop

## Files

| File | Description |
|------|-------------|
| `index.html` | Main app (light theme, Font Awesome icons) |
| `Ekart2/index.html` | Alternate app (dark/light theme, SVG icons) |
| `style.css` | Styles |
| `data.js` | Pincode dataset (11,764 entries) |
| `script.js` | Search & render logic |

## How to Use

1. Open `index.html` or `Ekart2/index.html` in any browser
2. Enter a 6-digit pincode
3. Press Enter or click Search
4. View serviceability details instantly

## Dataset

11,764 pincodes with branch, service type (STD/ODA), road & air service status, and active/inactive status.
