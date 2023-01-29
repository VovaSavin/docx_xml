import random
import datetime
import os

import xml.etree.ElementTree as Et
import zipfile

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from datas import (
    main_peoples,
    peoples,
)


class Extractor:
    """
    Extract xml file from docx file
    """

    def __init__(self, date_yesterday, date_before_yesterday):
        self.date_yesterday = date_yesterday
        self.date_before_yesterday = date_before_yesterday

    def extract_xml(
            self, zp_docx, yesterday, zp_docx_2, before_yesterday
    ):
        try:
            os.mkdir(f"./word_{yesterday}")
        except FileExistsError:
            print("File YESTERDAY exists. I do not create the file!")
        zp_docx.namelist()
        zp_docx.extractall(f"./word_{yesterday}")
        cur_xml = Et.parse(f"./word_{yesterday}/word/document.xml")
        cur_xml_read = cur_xml.getroot()
        # not_patrol = parse_xml(cur_xml_read)
        if before_yesterday is None and zp_docx_2 is None:
            print("Not_patrol")
            # print(not_patrol)

        else:
            zp_docx_2.namelist()
            zp_docx_2.extractall(f"./word_{before_yesterday}")
            before_cur_xml = Et.parse(f"./word_{before_yesterday}/word/document.xml")
            before_cur_xml_read = before_cur_xml.getroot()
            # before_not_patrol = parse_xml(before_cur_xml_read)
            print("Sum_not_patrol")

    def call_context_manager(self):
        """
        Extract files
        :return:
        """
        if self.date_yesterday and self.date_before_yesterday:
            with zipfile.ZipFile(
                    f"Patrol_{self.date_yesterday}.docx"
            ) as zp_docx, zipfile.ZipFile(
                f"Patrol_{self.date_before_yesterday}.docx"
            ) as zp_docx_2:
                self.extract_xml(zp_docx, self.date_yesterday, zp_docx_2, self.date_before_yesterday)
        elif self.date_yesterday:
            with zipfile.ZipFile(
                    f"Patrol_{self.date_yesterday}.docx"
            ) as zp_docx:
                self.extract_xml(zp_docx, self.date_yesterday, None, None)


class Existor:
    """
    Exist *.docx files
    """

    def __init__(self, date_yesterday, date_before_yesterday):
        self.date_yesterday = date_yesterday
        self.date_before_yesterday = date_before_yesterday
        self.if_path_before_yesterday = os.path.exists(f"./Patrol_{date_before_yesterday}.docx")
        self.if_path_yesterday = os.path.exists(f"./Patrol_{date_yesterday}.docx")

    def extract_data_from_docx(self, first_path=False, second_path=False):
        """
        Check path to file and call class for extract xml from docx
        :return:
        """
        if first_path and second_path:
            Extractor(self.date_yesterday, self.date_before_yesterday)
        elif first_path:
            Extractor(None, self.date_before_yesterday)
        elif second_path:
            Extractor(self.date_yesterday, None)
        else:
            pass


a = Existor(
    date_yesterday=datetime.date.today() - datetime.timedelta(days=1),
    date_before_yesterday=datetime.date.today() - datetime.timedelta(days=2)
)
