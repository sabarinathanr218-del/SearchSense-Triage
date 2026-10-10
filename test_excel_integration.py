import openpyxl
import urllib.request
import json

# Create test XLSX workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Search Terms Report"
ws.append(["Search term", "Match type", "Campaign", "Clicks", "Cost", "Conversions"])
ws.append(["buy dell XPS laptop", "Exact", "Laptop Sales", 25, 45.0, 2])
ws.append(["free laptop wallpaper", "Broad", "Laptop Sales", 10, 8.0, 0])
wb.save("test_organizer.xlsx")
print("Created test_organizer.xlsx successfully.")

# Upload file to local server
with open("test_organizer.xlsx", "rb") as f:
    content = f.read()

boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
header = (
    f"--{boundary}\r\n"
    'Content-Disposition: form-data; name="file"; filename="Search Terms Report.xlsx"\r\n'
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'
).encode("utf-8")
footer = f"\r\n--{boundary}--\r\n".encode("utf-8")

body = header + content + footer

req = urllib.request.Request("http://localhost:8505/api/upload", data=body, headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
resp = urllib.request.urlopen(req)
print("Upload API Response:", resp.read().decode("utf-8"))

# Test running analysis on uploaded Excel data
resp_an = urllib.request.urlopen("http://localhost:8505/api/analyze")
print("Analyze API Response:", resp_an.read().decode("utf-8"))

# Test state API
resp_st = urllib.request.urlopen("http://localhost:8505/api/state")
state = json.loads(resp_st.read().decode("utf-8"))
print("Decision counts on uploaded Excel file:", state["decision_counts"])
