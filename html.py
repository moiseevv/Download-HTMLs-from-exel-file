import requests
from openpyxl import load_workbook as lwb

wb = lwb("Кт.xlsx")
ws = wb.active
row_count = ws.max_row
print(row_count)

for i in range(3194, row_count + 1):
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

        f = open(saved_xml_file, 'wb')
        f.write(html)
        f.close()
    except AttributeError:
        pass
    except UnicodeEncodeError:
        pass
    except FileExistsError:
        pass
