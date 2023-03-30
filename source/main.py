# import docx2txt
import docx
import pandas as pd
my_path = "Word/Tu_vung_doi_chieu_GiaLai/Viet Ba Na OCR/"
paths = ["tu vung doi chieu p1 done.docx", "tu vung doi chieu p2 done.docx", "tu vung doi chieu p3 done.docx", "tu vung doi chieu p4.docx"]
i = 0
# language0_data = []
# language1_data = []
# word_type_data = []
for path in paths:
    file = docx.Document(my_path + path)
    language0 = []
    language1 = []
    word_type = []
    for table in file.tables:
        for row in table.rows:
            if len(row.cells) != 2:
                continue
            try:
                VN, type = row.cells[0].text.split('-')
                language0.append(VN.strip())
                word_type.append(type.strip())
                language1.append(row.cells[1].text.strip())
            except:
                pass
    path = "Excel/df" + str(i) + ".xlsx"
    data = {'language0': language0,
            'language1': language1,
            'word_type': word_type}            
    df = pd.DataFrame(data)
    df.to_excel(path, index = False)
    i += 1
