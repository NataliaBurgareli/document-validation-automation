import pandas as pd
from playwright.sync_api import sync_playwright

INPUT_FILE = "document_list.xlsx"
OUTPUT_FILE = "result.xlsx"

df = pd.read_excel(INPUT_FILE)

results = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    context = browser.new_context()

    # Login tab
    login_page = context.new_page()

    print("=" * 50)
    print("LOGIN MANUALLY IN THE FIRST TAB")
    print("PRESS ENTER WHEN FINISHED")
    print("=" * 50)

    input()

    # Processing tab
    query_page = context.new_page()

    for _, row in df.iterrows():

        case_id = str(row["Case_ID"])
        document_url = str(row["Repository_Link"])

        print(f"\nCASE: {case_id}")
        print(f"URL: {document_url}")

        try:

            query_page.goto(
                document_url,
                wait_until="domcontentloaded",
                timeout=60000
            )

            query_page.wait_for_timeout(3000)

            file_count = query_page.locator(
                "[data-document-id]"
            ).count()

            print(f"Documents found: {file_count}")

            results.append(file_count)

        except Exception as error:

            print(f"Error: {error}")

            results.append("ERROR")

    browser.close()

df["Document_Count"] = results

df.to_excel(
    OUTPUT_FILE,
    index=False
)

print("\nProcessing completed.")
print(f"Output file: {OUTPUT_FILE}")
