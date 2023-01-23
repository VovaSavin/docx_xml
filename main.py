import random
import datetime
import os

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


def iters_to_docx_tb(cnt: int, rec: list, peps: list, rng: int, to_stop: int, job: str, xml_file: list):
    for x in range(rng):
        main_p = random.choice(peps)
        if main_p not in xml_file:
            rec.append(
                (
                    x + cnt,
                    job,
                    main_p[0],
                    main_p[1]
                )
            )
        if len(rec) == to_stop:
            break


def parse_xml(root_of_xml) -> list:
    temp = []
    for x in range(13):
        temp_inner = []
        for y in range(3):
            temp_inner.append(root_of_xml[0][1][x + 3][y + 1][1][0][0].text.strip())
            print(root_of_xml[0][1][x + 3][y + 1][1][0][0].text)
        print("##" * 8)
        temp_inner = tuple(temp_inner)
        temp.append(temp_inner)
    return temp


def get_data_from_docx(days_ago: int):
    yesterday = datetime.date.today() - datetime.timedelta(days=days_ago)
    # today = datetime.date.today()
    if os.path.exists(f"C:/Users/Vovick/randomPatrol/Patrol{yesterday}.docx"):
        with zipfile.ZipFile(f"Patrol{yesterday}.docx") as zp_docx:
            os.mkdir(f"./word{yesterday}")
            zp_docx.namelist()
            zp_docx.extractall(f"C:/Users/Vovick/randomPatrol/word{yesterday}")

    cur_xml = et.parse(f"./word{yesterday}/word/document.xml")
    cur_xml_read = cur_xml.getroot()
    not_patrol = parse_xml(cur_xml_read)
    docs(not_patrol)


def docs(xml_file: list):
    document = Document()

    d = document.add_heading(f'Patrol tomorrow{datetime.date.today()}', 0)
    d.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    records = []
    iters_to_docx_tb(1, records, main_peoples, 2, 2, "Main", xml_file)

    iters_to_docx_tb(3, records, peoples, 12, 14, "Simple", xml_file)

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

get_data_from_docx(1)
