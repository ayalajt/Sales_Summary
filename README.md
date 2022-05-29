This Excel Information Extractor program takes in Excel sheets and parses through them to extract information the client needs. This extracted information is then formatted cleanly using HTML, which displays food, wine, beer, and liquor sales, as well as sales tax and total sales. This HTML page is then converted to a PDF and uploaded to Google Drive using Google's API. A front-end GUI is also included, created in tkinter, for ease of use. 

To use:
1. Run sales_summary.py and then upload as many excels as you want
2. The excel parsing results will be stored locally in the "results" folder (will be created in the same directory if it does not exist)
3. These results can then be uploaded to Google Drive if the client wants to, by choosing "Yes" when prompted to upload to Drive

For privacy reasons, client_secrets.json is not included even though it is required. Uploading to Google Drive also requires me to whitelist the account that wants to upload.