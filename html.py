import requests
from openpyxl import load_workbook as lwb
column_kadastr = 5          # Как правило 4 - я
column_link_on_html = 45    # Как правило 45 - я


wb = lwb("Кт.xlsx")
ws = wb.active
row_count = ws.max_row
print('Число строк в файле = ', row_count)

for i in range(2, row_count + 1):
    print("Строка", i)
    try:
        link_html = ws.cell(i, column_link_on_html).hyperlink.display  # 46 строка это ссылка на html
        print("link_html ", link_html)
        kadastr_num_file = ws.cell(i, column_kadastr).value
        kadastr_num_file = kadastr_num_file.replace(':', "_")
        print('Кадастровый номер ', kadastr_num_file)
        saved_xml_file = str(f'file_html\\{kadastr_num_file}.html')
        r = requests.get(link_html)
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

