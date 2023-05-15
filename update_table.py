from datetime import datetime

# Read the excel htm file
line_list = []
with open("Turnier2_excel.htm", "r", encoding="windows-1252") as file:
    for line in file:
        line_list.append(line)

header_lines = []
body_lines = []

# Save the header and body line into lists
parse_header = False
parse_body = False
for line in line_list:
    
    if "<head>" in line:
        parse_header = True
        continue

    if "</head>" in line:
        parse_header = False
        continue
    
    if "<div" in line:
        parse_body = True
    
    if "</div>" in line:
        body_lines.append(line)
        parse_body = False
    
    if parse_header:
        if '<style id="Turnier2_26858_Styles">' in line:
            line_clean = str(line).replace("<!--table", "")
            header_lines.append(line_clean)
        elif '{mso-displayed-decimal-separator:"\.";' in line:
            continue
        elif 'mso-displayed-thousand-separator' in line:
            continue
        elif '--></style>' in line:
            header_lines.append("</style>\n")
        else:
            header_lines.append(line)
    if parse_body:
        body_lines.append(line)
    
    


line_list = []
with open("index.html", "r", encoding="windows-1252") as file:
    for line in file:
        line_list.append(line)

new_line_list = []
skip = False
for line in line_list:

    if "<!-- Excel header start -->" in line:
        skip = True
        new_line_list.append(line)
        for new_line in header_lines:
            new_line_list.append(new_line)
    elif "<!-- Excel header end -->" in line:
        skip = False
        new_line_list.append(line)
    elif "<!-- Table start -->" in line:
        skip = True
        new_line_list.append(line)
        for new_line in body_lines:
            new_line_list.append(new_line)
    elif "<!-- Table end -->" in line:
        skip = False
        new_line_list.append(line)
    elif "Welcome back. Table last updated on: &nbsp" in line:
        datestr = datetime.now().strftime("%H:%M, %d.%m.%Y")
        print(datestr)
        new_line_list.append(f'<p class="text-white mt-5 mb-5">Welcome back. Table last updated on: &nbsp {datestr} </p>\n')
    elif not skip:
        new_line_list.append(line)



with open("index.html", "w", encoding="windows-1252") as file:
    for line in new_line_list:
        file.write(line)
