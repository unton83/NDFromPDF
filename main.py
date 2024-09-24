import os
import re
from curl_cffi import requests
import csv
import datetime as dt
from pypdf import PdfReader
import xlsxwriter
import ezdxf


patterns = [
    r'ГОСТ\s*Р?\s*\d+-\d{2,4}',             # ГОСТ
    r'(?i:СТО|ТУ)\s*(?:\d+-){1,4}\d{2,4}',  # СТО или ТУ
    r'(?i:Серия)\s*(?:\d+.){1,4}.\d+-\d+',  # Серия
    r'СП\s+(?:\d+\.){1,2}\d+(?:-\d+)?',     # СП
]


class WrongNDLabel(Exception):
    pass


class NDEntity:

    def __init__(self, label, full_name=None, path=None):
        self.label: str = label
        self.full_name: str = full_name if full_name else ''
        self.path: str = path if path else ''


class NDList:

    def __init__(self):
        self.list: list[NDEntity] = []

    def labels(self) -> list[str]:
        labels_list = []
        for entity in self.list:
            labels_list.append(entity.label)
        return labels_list

    def collect(self, dirpath):
        pdfs = os.listdir(dirpath)
        for pdf in pdfs:
            # print(pdf)
            pdf_file_fullname = f"{dirpath}\\{pdf}"
            reader = PdfReader(pdf_file_fullname)
            number_of_pages = len(reader.pages)
            # print(f"{number_of_pages = }")
            for page in range(number_of_pages):
                text = reader.pages[page].extract_text()
                # print(text)
                for pattern in patterns:
                    matches = re.findall(pattern, text)
                    for match in matches:
                        label = match.replace('\n', ' ').replace('  ', ' ')
                        if label not in self.labels():
                            self.list.append(NDEntity(label=label, path=f"{pdf}, лист {page+1}"))
                        else:
                            for ND in self.list:
                                if label == ND.label:
                                    ND.path += f"; {pdf}, лист {page+1}"
            # break

    @staticmethod
    def get_full_name(label: str):
        search_label = label.replace(' ', '+')
        url = f"https://docs.cntd.ru/api/search/intellectual/documents?q={search_label}"
        # print(url)

        payload = {}
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'ru,en;q=0.9'
        }

        response = requests.request("GET", url, headers=headers, data=payload)

        js = response.json()
        data = js['documents']['data']  # [0]['names']
        # if len(data) == 0:
        #     return f"    <— проверьте правильность записи!"
        for item in data:
            if item['names'][0].startswith(label):
                full_name = item['names'][0].removeprefix(label).strip()  # .removeprefix('"').removesuffix('"')
                if label.startswith("Серия"):
                    pattern = 'Выпуск'
                    match = re.search(pattern, full_name)
                    start_char = match.start()
                    full_name = full_name[:start_char]
                    full_name = full_name.strip().strip('.')
                return full_name
        raise WrongNDLabel

    def get_names(self):
        self.list.sort(key=lambda i: i.label, reverse=True)
        for ND in self.list:
            try:
                ND.full_name = self.get_full_name(ND.label)
            except WrongNDLabel:
                ND.full_name = f"Ошибка"
            except Exception as ex:
                ND.full_name = ex.__str__()

    # def write_csv(ND_dic: dict):
    #     file_name = dt.datetime.now().strftime("%Y%m%d_%H%M") + ".csv"
    #     with open(file_name, 'w', newline='', encoding='utf-8') as file:
    #         fieldnames = ['label', 'full_name']
    #         writer = csv.writer(file, quotechar='"', delimiter=';')
    #         for key, value in ND_dic.items():
    #             writer.writerow([key, value])

    def write_xlsx(self):
        file_name = dt.datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx"

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet()
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 1, 95)
        # Start from the first cell. Rows and columns are zero indexed.
        row = 0

        # Iterate over the data and write it out row by row.
        for ND in self.list:
            worksheet.write(row, 0, ND.label)
            worksheet.write(row, 1, ND.full_name)
            worksheet.write(row, 2, ND.path)
            row += 1
        workbook.close()

    # def write_dxf(ND_dict: dict):
    #     file_name = dt.datetime.now().strftime("%Y%m%d_%H%M") + ".dxf"
    #
    #     def get_mat_symbol():
    #         p1 = 0.5
    #         p2 = 0.25
    #         points = [(p1, p2), (p2, p1), (-p2, p1), (-p1, p2), (-p1, -p2),
    #                   (-p2, -p1), (p2, -p1), (p1, -p2)]
    #         polygon = dxf.polyline(points, color=2)
    #         polygon.close()
    #         attdef = dxf.attdef(text='0', tag='num', height=0.7, color=1,
    #                             halign=dxfwrite.CENTER, valign=dxfwrite.MIDDLE
    #                             )
    #         symbolblock = dxf.block('matsymbol')
    #         symbolblock.add(polygon)
    #         symbolblock.add(attdef)
    #         dwg.blocks.add(symbolblock)
    #         return symbolblock
    #
    #     dwg = dxf.drawing(file_name)  # create a drawing
    #     rows = len(ND_dict.keys()) + 1
    #     table = dxf.table(insert=(0, 0), nrows=rows, ncols=3)
    #     # create a new styles
    #     ctext = table.new_cell_style('ctext', textcolor=7, textheight=2.5,
    #                                  halign=dxfwrite.LEFT,
    #                                  valign=dxfwrite.MIDDLE
    #                                  )
    #     # modify border settings
    #     border = table.new_border_style(color=6, linetype='DOT', priority=51)
    #     ctext.set_border_style(border, right=False)
    #
    #     # table.new_cell_style('vtext', textcolor=3, textheight=0.3,
    #     #                      rotation=90,  # vertical written
    #     #                      halign=dxfwrite.CENTER,
    #     #                      valign=dxfwrite.MIDDLE,
    #     #                      bgcolor=8,
    #     #                      )
    #     row = 0
    #     # col = 0
    #     for key, value in ND_dict.items():
    #         col = 0
    #         table.set_col_width(col, 60)
    #         table.set_row_height(row, 8)
    #         table.text_cell(row, col, key, style='ctext')
    #         col = col + 1
    #         table.set_col_width(col, 95)
    #         table.text_cell(row, col, f"{value}\n", style='ctext')
    #         col = col + 1
    #         table.set_col_width(col, 30)
    #         # table.text_cell(row, col, value, style='ctext')
    #         row += 1
    #
    #     # # set colum width, first column has index 0
    #     # table.set_col_width(1, 7)
    #     #
    #     # # set row height, first row has index 0
    #     # table.set_row_height(1, 7)
    #     #
    #     # # create a text cell with the default style
    #     # cell1 = table.text_cell(0, 0, 'Zeile1\nZeile2', style='ctext')
    #     #
    #     # # cell spans over 2 rows and 2 cols
    #     # cell1.span=(2, 2)
    #     #
    #     # cell2 = table.text_cell(4, 0, 'VERTICAL\nTEXT', style='vtext', span=(4, 1))
    #     #
    #     # # create frames
    #     # table.frame(0, 0, 10, 2, 'framestyle')
    #     #
    #     # # because style is defined by a namestring
    #     # # style can be defined later
    #     # hborder = table.new_border_style(color=4)
    #     # vborder = table.new_border_style(color=17)
    #     # table.new_cell_style('framestyle', left=hborder, right=hborder,
    #     #                      top=vborder, bottom=vborder)
    #     # mat_symbol = get_mat_symbol()
    #     #
    #     # table.new_cell_style('matsym',
    #     #                      halign=dxfwrite.CENTER,
    #     #                      valign=dxfwrite.MIDDLE,
    #     #                      xscale=0.6, yscale=0.6)
    #     #
    #     # # add table as anonymous block
    #     # # dxf creation is only done on save, so all additional table inserts
    #     # # which will be done later, also appear in the anonymous block.
    #     #
    #     # dwg.add_anonymous_block(table, insert=(40, 20))
    #     #
    #     # # if you want different tables, you have to deepcopy the table
    #     # newtable = deepcopy(table)
    #     # newtable.new_cell_style('57deg', textcolor=2, textheight=0.5,
    #     #                      rotation=57, # write
    #     #                      halign=dxfwrite.CENTER,
    #     #                      valign=dxfwrite.MIDDLE,
    #     #                      bgcolor=123,
    #     #                      )
    #     # newtable.text_cell(6, 3, "line one\nline two\nand line three",
    #     #                    span=(3,3), style='57deg')
    #     # dwg.add_anonymous_block(newtable, basepoint=(0, 0), insert=(80, 20))
    #     #
    #     # # a stacked text: Letters are stacked top-to-bottom, but not rotated
    #     # table.new_cell_style('stacked', textcolor=6, textheight=0.25,
    #     #                      halign=dxfwrite.CENTER,
    #     #                      valign=dxfwrite.MIDDLE,
    #     #                      stacked=True)
    #     # table.text_cell(6, 3, "STACKED FIELD", span=(7, 1), style='stacked')
    #     #
    #     # for pos in [3, 4, 5, 6]:
    #     #     blockcell = table.block_cell(pos, 1, mat_symbol,
    #     #                                 attribs={'num': pos},
    #     #                                 style='matsym')
    #
    #     dwg.add(table)
    #     dwg.save()
    #     print("drawing '%s' created.\n" %  file_name)

    def write_dxf(self):
        file_name = dt.datetime.now().strftime("%Y%m%d_%H%M") + ".dxf"

        # Create a new DXF document
        doc = ezdxf.new(dxfversion='R2018')
        msp = doc.modelspace()

        # Define table parameters
        start_x = 0
        start_y = 0
        # cell_width = 20
        cell_height = 8
        rows = len(self.list) + 1

        col_1 = 60
        col_2 = 95
        col_3 = 30

        attribs_lable = {
            "char_height": 2.5,
            "width": 50,
            "style": "Arial",
        }
        attribs_name = {
            "char_height": 2.5,
            "width": 92,
            "style": "Arial",
            "line_spacing_factor": 0.96,
        }

        # Draw the table
        for row in range(rows + 1):
            msp.add_line([start_x, start_y + row * cell_height], [start_x + (col_1+col_2+col_3), start_y + row * cell_height])

        # for col in range(cols + 1):
        msp.add_line([start_x, start_y], [start_x, start_y + rows * cell_height])
        msp.add_line([start_x + col_1, start_y], [start_x + col_1, start_y + rows * cell_height])
        msp.add_line([start_x + col_1 + col_2, start_y], [start_x + col_1 + col_2, start_y + rows * cell_height])
        msp.add_line([start_x + col_1 + col_2 + col_3, start_y], [start_x + col_1 + col_2 + col_3, start_y + rows * cell_height])

        # Optionally, add text to the cells
        for row, ND in enumerate(self.list):
            if 'Ошибка' in ND.full_name:
                full_name = f"{ND.full_name}: {ND.path}"
                attribs_name.update({'color': 1})
            else:
                full_name = ND.full_name
                attribs_name.update({'color': 7})
            msp.add_mtext(ND.label, attribs_lable).set_location(insert=(start_x + 2, start_y + row * cell_height + 5.5))
            msp.add_mtext(full_name, attribs_name).set_location(insert=(start_x + col_1 + 2, start_y + row * cell_height + 7))

        # Save the DXF document
        doc.saveas(file_name)


def main(dirpath):
    ND_list = NDList()
    ND_list.collect(dirpath)
    ND_list.get_names()
    # write_csv(ND_dict)
    ND_list.write_xlsx()
    ND_list.write_dxf()


if __name__ == '__main__':
    cur_dir = os.getcwd()
    pdf_dir = cur_dir + "\pdfs"
    # print(pdf_dir)
    main(pdf_dir)
