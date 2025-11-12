# Enable Sheets API and Charts API for Summary Sheet

The Summary sheet builder uses the Google Sheets API to create slicers and charts programmatically. You need to enable these APIs in your Apps Script project.

## Step 1: Enable Advanced Google Services

1. Open your Apps Script project:
   - Go to https://script.google.com
   - Open your project (ID: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`)

2. Enable Google Sheets API:
   - Click **Resources** → **Advanced Google services**
   - Find **Google Sheets API** in the list
   - Toggle it **ON**
   - Click **OK**

3. Enable in Google Cloud Console:
   - Click the link next to "Google Sheets API" that says "Google Cloud Platform API Dashboard"
   - This opens the Google Cloud Console
   - Make sure the API is enabled (should show "API enabled")
   - If not enabled, click **Enable**

## Step 2: Verify API Access

The code will automatically check if the Sheets API is available. If it's not enabled:
- Slicers will be created as placeholder cells (you can convert them manually)
- Chart data will be created, but the chart will need to be created manually

## Step 3: Run Build Summary Sheet

1. In your Google Sheet, go to: **🤖 Bob Automation > Build Summary Sheet**
2. The function will:
   - Build the data table
   - Create slicers (if API enabled)
   - Create chart (if API enabled)
   - Apply formatting

## Troubleshooting

### Error: "Sheets is not defined"
- **Solution**: Enable Google Sheets API (see Step 1)

### Slicers not appearing
- **Solution**: Check if Sheets API is enabled. If not, placeholder cells will be created that you can convert to slicers manually:
  1. Select the data range
  2. Go to **Data > Create a filter**
  3. Right-click on filter icons → **Add a slicer**

### Chart not appearing
- **Solution**: Chart data is always created. If the chart doesn't appear:
  1. Select the chart data (columns F-H, rows 1-7)
  2. Go to **Insert > Chart**
  3. Choose **Column chart**
  4. Customize colors as needed

## Manual Slicer Creation (Fallback)

If API is not enabled, you can create slicers manually:

1. Select your data table (starting at row 20)
2. Go to **Data > Create a filter**
3. For each slicer:
   - Click the filter icon on the column
   - Click **Add a slicer**
   - Position it in the slicer area (rows 4-15)
   - Customize color and title

## Notes

- The function works even without the API enabled (uses fallback methods)
- API enables full automation of slicers and charts
- All data and formulas are created regardless of API status

