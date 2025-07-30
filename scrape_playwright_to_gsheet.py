from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
import re

# ✅ CONFIG
SHEET_ID = "15kevgDAKABBLRG_HgVTE-LoF0Q8i6rvFefIk6Bp2b8A"
CREDENTIALS_FILE = "credentials.json"

ALL_CAREER_PATHS = {
    "business-analytics": [
        "Microsoft 365 Copilot for Business Professionals",
        "Microsoft Excel Advanced",
        "Microsoft Excel Power Query",
        "Power BI Desktop for Business Analytics",
        "Data Analysis Expression (DAX) for Power BI"
    ],
    "prompt-engineer": [
        "Python Programming",
        "Machine Learning using Python",
        "Generative AI for Business Transformation",
        "AI Agents with Microsoft Copilot Studio",
        "AI Automation Agent with Make.com",
        "Workflow Automation with n8n"
    ],
    "data-analyst": [
        "Power BI Desktop for Business Analytics",
        "Power BI Advanced Power Query",
        "Data Analysis Expression (DAX) for Power BI",
        "Power BI Advanced Visualization and AI",
        "Data Model for Power BI",
        "Generative AI for Business Transformation"
    ],
    "business-intelligence-development": [
        "Microsoft SQL Server Essential",
        "ETL with SQL Server Integration Service (SSIS)",
        "Microsoft Fabric Essential for Business",
        "Power BI Desktop for Business Analytics",
        "Data Model for Power BI"
    ],
    "citizen-developer": [
        "Power Apps for Business",
        "Advanced Power Apps for Business",
        "Power Automate (Cloud) for Business Automation",
        "Advanced Power Automate (Cloud)",
        "AI Builder in Power Platform for Business"
    ],
    "rpa-developer": [
        "UiPath for Business Automation",
        "Power Automate (Desktop) for Business Automation",
        "Advanced Power Automate (Desktop)",
        "Generative AI for Business Transformation",
        "AI Automation Agent with Make.com",
        "Workflow Automation with n8n"
    ],
    "power-automate-specialist": [
        "Power Automate (Cloud) for Business Automation",
        "Advanced Power Automate (Cloud)",
        "Power Automate (Desktop) for Business Automation",
        "Advanced Power Automate (Desktop)",
        "AI Builder in Power Platform for Business"
    ],
    "web-development": [
        "Programming in C# with Visual Studio",
        "Querying Data with T-SQL",
        "ASP.NET Core MVC",
        "ASP.NET Core Web API & Security",
        "Github Copilot"
    ],
    "accounting-and-finance": [
        "Microsoft Excel Intermediate",
        "Microsoft Excel Advanced",
        "Power BI Desktop for Business Analytics",
        "Microsoft 365 Copilot for Business Professionals"
    ],
    "visual-communication-and-presentation": [
        "Infographics & Digital Media with Advanced Microsoft PowerPoint",
        "Canva Pro for Smart Working",
        "Generative AI for Business Transformation",
        "Microsoft Excel Advanced PivotTable and PivotChart",
        "Power BI Desktop for Business Analytics"
    ]
}

# 📥 STEP 1: SCRAPE
def get_course_rounds(course_list):
    print("🧪 DEBUG: เริ่มดึงข้อมูลคอร์ส")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
    try:
        page.goto("https://www.9experttraining.com/schedule", timeout=90000)
        page.wait_for_selector("table", timeout=90000)
    except Exception as e:
        page.screenshot(path="debug.png")
        print("Error occurred, screenshot saved.")
        raise e

        tables = page.query_selector_all("table")
        course_data = {}

        for table in tables:
            # ✅ หาเดือนจาก thead
            header_cells = table.query_selector_all("thead tr th")
            month_list = []
            for th in header_cells:
                text = th.inner_text().strip()
                if re.match(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$", text):
                    month_list.append(text)
            if not month_list:
                continue

            # ✅ อ่าน tbody → คอร์ส → รอบอบรม
            rows = table.query_selector_all("tbody tr")
            for row in rows:
                cells = row.query_selector_all("td")
                if len(cells) < 2:
                    continue

                course_name = cells[1].inner_text().strip()
                for target in course_list:
                    if course_name == target:
                        print(f"🎯 พบคอร์สตรงกัน: {course_name}")
                        rounds = []
                        types = []

                        for i, cell in enumerate(cells[4:4+len(month_list)]):
                            month = month_list[i] if i < len(month_list) else "?"

                            lis = cell.query_selector_all("li")
                            for li in lis:
                                a_tags = li.query_selector_all("a")
                                for a_tag in a_tags:
                                    date_text = a_tag.inner_text().strip()
                                    if not re.match(r"^\d{2}(\s*-\s*\d{2})?$", date_text):
                                        continue

                                    type_text = "-"
                                    img = li.query_selector("img")
                                    if img:
                                        alt = img.get_attribute("alt") or ""
                                        if "hybrid" in alt.lower():
                                            type_text = "Hybrid"
                                        elif "class room" in alt.lower():
                                            type_text = "Class Room"

                                    full_round = f"{date_text} {month} ({type_text})"
                                    rounds.append(full_round)
                                    types.append(type_text)

                        course_data[target] = {"rounds": rounds, "types": types}

        browser.close()
        return course_data

# 📤 STEP 2: WRITE TO GOOGLE SHEET
def update_google_sheet(sheet_name, course_list, data):
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SHEET_ID).worksheet(sheet_name)

    sheet.clear()
    max_rounds = max((len(d["rounds"]) for d in data.values()), default=0)

    header = ["Course Name"] + [f"Round {i+1}" for i in range(max_rounds)]
    sheet.append_row(header)

    for course in course_list:
        rounds = data.get(course, {}).get("rounds", [])
        types = data.get(course, {}).get("types", [])
        sheet.append_row([course] + rounds)
        print(f"📝 เขียนข้อมูล: {course} → {rounds}")

# ▶️ MAIN
if __name__ == "__main__":
    for sheet_name, course_list in ALL_CAREER_PATHS.items():
        print(f"\n📌 กำลังดึงข้อมูล Career Path: {sheet_name}")
        data = get_course_rounds(course_list)
        update_google_sheet(sheet_name, course_list, data)

    print("\n✅ เสร็จสิ้นการดึงข้อมูลทั้งหมดเรียบร้อยแล้ว")
