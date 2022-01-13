import pdfplumber
import re
import pandas as pd
import os

def parse(pdf_file, starting_page):
    # Extract data from pdf file
    print(f'\n\n----- EXTRACTING DATA FROM: {pdf_file} -----\n\n')
    with pdfplumber.open(pdf_file) as pdf:
        pages = []

        for page in pdf.pages[starting_page:]:
            print(page.extract_text())
            lines = page.extract_text().split('\n')
            for line in lines:
                if ' ID ' in line:
                    index = lines.index(line)
                    for i in range(5):
                        lines.insert(i, lines.pop(index + i))
                    break
            page = '\n'.join(lines)
            pages.append(page)

    # Parsing extracted data
    print('\n\n----- PARSING EXTRACTED DATA -----\n\n')
    out_data = []

    emps = re.split(r' ID ', '\n'.join(pages))[:]
    for (index, emp) in zip(range(len(emps) - 1), emps[1:]):
        name = emps[index].split()[-1].replace(',', ' ')
        id = emps[index + 1].split()[0]
        term_date = emps[index + 1].split()[2]
        try:
            int(term_date.split('/')[0])
        except:
            term_date = 'None'
        department = emps[index + 1].split('\n')[1].split()[2]
        hire_date = emps[index + 1].split('\n')[4].split()[1]

        year_data = re.split(r'YTD \d', emp)
        for (y_index, y_data) in zip(range(1, len(year_data)), year_data):
            year = year_data[y_index].split('\n')[1].split()[0]
            
            qtr_data = re.split(r'QTR', y_data)
            for (q_index, q_data) in zip(range(1, len(qtr_data)), qtr_data):
                qtr = qtr_data[q_index].split()[0]
                qtr_year = 'Q' + qtr + '-' + year

                dates = re.compile(r'\d{2}/\d{2} ').findall(q_data)
                date_split = re.split(r'\d{2}/\d{2} ', q_data)[1:]
                for (date, data) in zip(dates, date_split):
                    reg_hrs, reg_pay, ot_hrs, ot_pay, dt_hrs, dt_pay, sick_hrs, sick_pay, hol_hrs, hol_pay, bonus, other, comm = ('None' for _ in range(13))
                    for row in data.split('\n'):
                        try:
                            row = row.split()
                            if row[0] == 'Reg':
                                reg_hrs, reg_pay = row[1:3]
                                other = row[-2]
                            elif row[0] == 'Overtime':
                                ot_hrs, ot_pay = row[1:3]
                            elif row[0] == 'Doubletime':
                                dt_hrs, dt_pay = row[1:3]
                            elif row[0] == 'Sick':
                                sick_hrs, sick_pay = row[1:3]
                            elif row[0] == 'Holiday':
                                hol_hrs, hol_pay = row[1:3]
                            elif row[0] == 'Commission':
                                comm = row[1]
                            elif row[0] == 'Bonus':
                                bonus = row[1]
                        except:
                            continue

                    data_dict = {'EMP ID': id, 'EMP Name': name, 'EMP Department': department, 'EMP Hire date': hire_date, 'EMP Term date': term_date, 'Date': date, 'QTR-Year':qtr_year, 'Reg Hrs': reg_hrs, 'OT Hrs': ot_hrs, 'DT Hrs': dt_hrs, 'Sick Hrs': sick_hrs, 'Hol Hrs': hol_hrs, 'Reg Pay': reg_pay, 'OT Pay': ot_pay, 'DT Pay': dt_pay, 'Sick Pay': sick_pay, 'Hol Pay': hol_pay, 'Comm': comm, 'Bonus': bonus, 'Other': other}
                    print(data_dict)
                    out_data.append(data_dict)
    
    return out_data

# Output result
out_data = parse('Paychex history sales.pdf', 8) + parse('Paychex History Service.pdf', 1)
df = pd.DataFrame(out_data)
OUT_FILE = 'out_data.xlsx'
if os.path.exists(OUT_FILE):
    os.remove(OUT_FILE)
df.to_excel(OUT_FILE, index=False)
