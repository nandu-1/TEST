from pptx import Presentation
from pptx.util import Inches, Pt

# Create a presentation object
prs = Presentation()

# Define slide dimensions (optional)
prs.slide_width = Inches(25)
prs.slide_height = Inches(12)

# Add a blank slide layout
slide_layout = prs.slide_layouts[6]  # Blank layout
slide = prs.slides.add_slide(slide_layout)

# Add the title 1 inch from the left and top, and center it
title_text = "Company name (prepared on today's date)"
title_shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(1))
title_frame = title_shape.text_frame
title_frame.text = title_text
title_frame.paragraphs[0].alignment = 1  # Center alignment
title_frame.paragraphs[0].font.size = Pt(22)

# Data for the tables
company_details = [
    ["Company Details"],
    ["WEBSITE"],
    ["HQ ADDRESS"],
    ["ANNUAL GROWTH"],
    ["# OF EMPLOYEES"],
    ["INDUSTRY"],
]

key_contacts = [
    ["Key Contacts", "", "", ""],  # Padded with empty strings
    ["CEO/Founder", "Nandana\nph no:23234234\nmail:ganuradhakrish@gmail.com", "VP, Finance", "Kartikeya\nph no:23234234\nmail:ganuradhakrish@gmail.com"],
    ["Co-CEO/Co-Founder", "Nandana\nph no:23234234\nmail:ganuradhakrish@gmail.com", "VP, Sales", "Gannuji\nph no:23234234\nmail:ganuradhakrish@gmail.com"],
]

rv_info = [
    ["RV Info",""],
    ["WCIS_ID", "Coverage Leads"],
    ["WCIS_SUP_ID", "Last call reported by"],
    ["WCIS_STATUS", "Last call reported on"],
    ["BDO", "Call Logs"],
    ["RM", "Similar companies"],
]

potential_banking_relationship = [[]]

noteworthy_information = [[]]

# Ensure all sublists in the data are the same length
max_columns_company = max(len(row) for row in company_details)
max_columns_key = max(len(row) for row in key_contacts)
max_columns_rv = max(len(row) for row in rv_info)
max_columns_potential = max(len(row) for row in potential_banking_relationship)
max_columns_noteworthy = max(len(row)for row in noteworthy_information)

# Verify the data structures have uniform row lengths
for idx, row in enumerate(company_details):
    if len(row) != max_columns_company:
        raise ValueError(f"Row {idx} in 'company_details' has inconsistent column length.")

for idx, row in enumerate(key_contacts):
    if len(row) != max_columns_key:
        raise ValueError(f"Row {idx} in 'key_contacts' has inconsistent column length.")

for idx, row in enumerate(rv_info):
    if len(row) != max_columns_rv:
        raise ValueError(f"Row {idx} in 'rv_info' has inconsistent column length.")

# Define the width for each table
table_width = Inches(13)
left_margin = Inches(1)
right_margin = Inches(15)  # Right side starts after the first three tables

# Add the first table (Company Details)
top_margin = Inches(2)  # 1 inch below the title
table1_height = Inches(1.5)  # Adjust as needed
table1_top = top_margin
table1 = slide.shapes.add_table(len(company_details), max_columns_company, left_margin, table1_top, table_width, table1_height).table
for row_idx, row in enumerate(company_details):
    for col_idx, cell_text in enumerate(row):
        cell = table1.cell(row_idx, col_idx)
        cell.text = cell_text
        cell.text_frame.word_wrap = True  # Enable text wrapping
table1.rows[0].height = Inches(0.25)

# Add the second table (Key Contacts)
table2_top = table1_top + table1_height + Inches(1)
table2_height = Inches(2)  # Adjust as needed
table2 = slide.shapes.add_table(len(key_contacts), max_columns_key, left_margin, table2_top, table_width, table2_height).table
for row_idx, row in enumerate(key_contacts):
    for col_idx, cell_text in enumerate(row):
        cell = table2.cell(row_idx, col_idx)
        cell.text = cell_text
        cell.text_frame.word_wrap = True  # Enable text wrapping
table2.rows[0].height = Inches(0.25)

# Add the third table (RV Info)
table3_top = table2_top + table2_height + Inches(1)
table3_height = Inches(4)  # Adjust as needed
table3 = slide.shapes.add_table(len(rv_info), max_columns_rv, left_margin, table3_top, table_width, table3_height).table
for row_idx, row in enumerate(rv_info):
    for col_idx, cell_text in enumerate(row):
        cell = table3.cell(row_idx, col_idx)
        cell.text = cell_text
        cell.text_frame.word_wrap = True  # Enable text wrapping
table3.rows[0].height = Inches(0.25)

# Add the fourth table (to the right of the first table)
table4_top = table1_top  # Align with the first table's top
table4_height = Inches(7.5)  # Adjust as needed
table_width = prs.slide_width-right_margin-Inches(1)
table4 = slide.shapes.add_table(2, 1, right_margin, table4_top, table_width, table4_height).table

header_cell = table4.cell(0, 0)
header_cell.text = "Potential Banking Relationship"

# Apply formatting to the header cell
header_cell.text_frame.paragraphs[0].font.size = Pt(18)  # Font size
header_cell.text_frame.paragraphs[0].alignment = 1  # Center alignment
header_cell.text_frame.paragraphs[0].font.bold = True
table4.rows[0].height = Inches(0.25)

for row_idx, row in enumerate(potential_banking_relationship):
    for col_idx, cell_text in enumerate(row):
        cell = table4.cell(row_idx, col_idx)
        cell.text = cell_text
        cell.text_frame.word_wrap = True  # Enable text wrapping

# Add the fifth table (to the right of the second table, aligned with the third table)
table5_top = table4_height-Inches(2)  # Align with the top of the third table
table5_height = Inches(10.5)  # Adjust as needed
table5 = slide.shapes.add_table(2, 1, right_margin, table5_top, table_width, table5_height).table

header_cell = table5.cell(0, 0)
header_cell.text = "Noteworthy Information"

# Apply formatting to the header cell
header_cell.text_frame.paragraphs[0].font.size = Pt(18)  # Font size
header_cell.text_frame.paragraphs[0].alignment = 1  # Center alignment
header_cell.text_frame.paragraphs[0].font.bold = True
table5.rows[0].height = Inches(0.25)

for row_idx, row in enumerate(noteworthy_information):
    for col_idx, cell_text in enumerate(row):
        cell = table5.cell(row_idx, col_idx)
        cell.text = cell_text
        cell.text_frame.word_wrap = True  # Enable text wrapping

# Save the presentation
prs.save("final_layout_presentation1.pptx")
