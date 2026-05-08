# Image Scoring Comparator & Auditor

A professional tool for auditing image assignment changes between two datasets. Designed for visual verification of high-volume item catalogs.

## 🚀 Getting Started

1. **Upload Baseline (V1):** Drag or select your previous version of the image assignment Excel file.
2. **Upload Target (Scored):** Drag or select the new file containing the results from the image scoring engine.
3. **Audit Results:**
   - **Visual Comparison:** Side-by-side view of Previous vs. Current images.
   - **Change Tracking:** Automatic detection of changes in Image Name, URL, Status, or Package.
   - **Deep Dive:** Click any row to open a full-screen preview for detailed inspection.
4. **Export:** Filter for "Changed Only" and export the results to a new Excel file for reporting or implementation.

## 📋 Expected File Format

The application expects Excel (.xlsx) or CSV files with the following columns:
- `TABLE_ID`
- `CSOR_IMAGE_NAME`
- `UPC`
- `CUSTOMER_ITEM_CODE`
- `GL`
- `SOURCE_IMAGE_URL`
- `DOWNLOAD_SOURCE`
- `IMAGE_STATUS`
- `PKG_NAME`

## 🔒 Security & Privacy

- **Local Processing:** All file processing happens directly in your browser.
- **No API Keys:** This application does not require any external API keys (Gemini, Maps, etc.) to function.
- **Data Privacy:** Your Excel data never leaves your machine. Images are loaded directly from the provided URLs via the browser.

## 💡 Pro Tips

- **Sorting:** Click on column headers (e.g., UPC, Baseline Assignment) to change the sorting order.
- **Filtering:** Use the "Quick UPC find" search bar to isolate specific items instantly.
- **Audit View:** Use the "Changed Only" toggle to quickly see where the scoring engine has made updates compared to your baseline.
