import pdfplumber
import pandas as pd

# Open the PDF file
with pdfplumber.open(r"C:\Users\MatthewHieger\Documents\Ad Hoc\One Medical Final Addendum.pdf") as pdf:
    # Extract data from the first page
    first_page = pdf.pages[4]
    
    # Extract table data from the page
    table = first_page.extract_table()

# Convert the table data to a DataFrame
df = pd.DataFrame(table[2:], columns=table[0])

print(df)

# Save the DataFrame to an Excel file
df.to_csv(r"C:\Users\MatthewHieger\Documents\Ad Hoc\One Medical Final Addendum6.csv", index=False)
