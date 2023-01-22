import random
import datetime

import xml.etree.ElementTree as et
import zipfile

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from datas import (
    main_peoples,
    peoples,
    today_people,
    today_main_people,
    today_patrol,
)


def iters_to_docx_tb(rec: list, peps: list, rng: int, to_stop: int, job: str):
    for x in range(rng):
        main_p = random.choice(peps)
        rec.append(
            (
                x + 1,
                job,
                main_p[0],
                main_p[1]
            )
        )
        if len(rec) == to_stop:
            break


def get_data_from_docx():
    with zipfile.ZipFile("Patrol2023-01-22.docx") as z:
        z.namelist()
        z.extractall()
    tr = et.parse("[Content_Types].xml")
    rr = tr.getroot()
    finder_xml = None
    for r in rr:
        if list(r.attrib.values())[0] == "/word/document.xml":
            finder_xml = "/word/document.xml"

    cur_xml = et.parse("word/document.xml")
    cur_xml_read = cur_xml.getroot()
    print(cur_xml_read[0][1][3][1][1][0][0].text)
    print(cur_xml_read[0][1][3][2][1][0][0].text)
    print(cur_xml_read[0][1][3][3][1][0][0].text)


def docs():
    document = Document()

    d = document.add_heading('Patrol tomorrow', 0)
    d.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    records = []
    iters_to_docx_tb(records, main_peoples, 2, 2, "Main")

    iters_to_docx_tb(records, peoples, 12, 14, "Simple")

    table = document.add_table(rows=1, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '№'
    hdr_cells[1].text = 'Job'
    hdr_cells[2].text = 'Name'
    hdr_cells[3].text = 'NumberAK'
    for num, qty, ids, desc in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(num)
        row_cells[1].text = str(qty)
        row_cells[2].text = ids
        row_cells[3].text = desc

    document.add_page_break()

    document.save(f'Patrol{datetime.datetime.today().date()}.docx')


#
# docs()
get_data_from_docx()
