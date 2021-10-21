import requests
from openpyxl import load_workbook as lwb


wb = lwb("Кт.xlsx")
ws = wb.active
row_count = ws.max_row
print(row_count)

for i in range(3194, row_count+1):
    print("Строка", i)
    try:
        link_xml = ws.cell(i, 45).hyperlink.display  # 46 строка это ссылка на xml
        print("link_xml ", link_xml)
        kadastr_num_file = ws.cell(i, 4).value
        kadastr_num_file = kadastr_num_file.replace(':', "_")
        print('Кадастровый номер ', kadastr_num_file)
        saved_xml_file = str(f'file_html\\{kadastr_num_file}.html')
        r = requests.get(link_xml)
        html = r.content
        #html = r.content
        # html = html.replace('\u04b7','')
        # html = html.replace('\u1d34','')
        # html = html.replace('\u1d7e','')
        # html = html.replace('\u04be','')
        # html = html.replace('\u04f1','')
        # html = html.replace('\u1d2d','')
        # html = html.replace('\xc1','')
        # html = html.replace('\u0469','')
        # html = html.replace('\u0477','')
        # html = html.replace('\u0485','')
        # html = html.replace('\u02cc','')
        # html = html.replace('\u2206','')
        # html = html.replace('\u02c2','')
        # html = html.replace('\u0512','')
        # html = html.replace('\u202a','')
        # html = html.replace('\xed','')
        # html = html.replace('\u04a3','')
        # html = html.replace('\u04aa','')
        # html = html.replace('\u049b','')
        # html = html.replace('\u02eb','')
        # html = html.replace('\u1d71','')
        # html = html.replace('\u1d24','')
        # html = html.replace('\u1d3b','')
        

        f = open(saved_xml_file, 'wb')
        f.write(html)
        f.close()



    except AttributeError:
        pass
    except UnicodeEncodeError:
        pass
    except FileExistsError:
        pass