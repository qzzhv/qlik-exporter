import time
from pathlib import Path
from uuid import uuid1

from loguru import logger
from win32com.client import Dispatch


class Exporter:
    _app = None
    doc_num = 0

    def __init__(self, qlik_path_file: str,
                 qlik_object_id: str,
                 qlik_export_path: str = Path(uuid1().hex + ".csv"),
                 qlik_bookmark: str = None,
                 qlik_fields: dict = None):

        """
        Export the sheet element specified by ID to CSV format. The class uses the installed Qlikview program.
        :param qlik_path_file: Address or path to the qvw file, like a "C:\123.qvw" or "qvp://qlivkiewserver/report.qvw"
        :param qlik_object_id: ID of the sheet object (in properties) like a "CH08" or "CH26-197"
        :param qlik_export_path: path for exporting the object to a csv file.
        :param qlik_bookmark: Name of the bookmark, if you want to apply it.
        :param qlik_fields: Fields for additional selection in [{"name": "FIELDNAME", "values": ["FIELDVALUE", ...]}, ...] format.
        """
        self._doc = None

        self.qlik_path_file = qlik_path_file

        prefix = "Server\\" if "-" in qlik_object_id else "Document\\"
        self.qlik_object_id = prefix + qlik_object_id

        self.qlik_export_path = qlik_export_path
        self.bookmark = qlik_bookmark
        self.qlik_fields = qlik_fields

    @property
    def app(self):
        Exporter._app = Dispatch("QlikTech.QlikView")
        return Exporter._app

    @property
    def doc(self):
        if self._doc is None:
            logger.debug(f"Open {self.qlik_path_file}")
            self._doc = self.app.OpenDoc(self.qlik_path_file)
            Exporter.doc_num += 1
        return self._doc

    def _runner(self):
        logger.debug(f"Selecting bookmark {self.bookmark}...")
        if self.bookmark is not None:
            self.doc.RecallUserBookmark(self.bookmark)
            self.doc.RecallDocBookmark(self.bookmark)
        logger.debug(f"Getting object {self.qlik_object_id}...")
        chart = self.doc.GetSheetObject(self.qlik_object_id)

        if self.qlik_fields is not None:
            for field in self.qlik_fields:
                for value in field["values"]:
                    logger.debug(f"Adding selection [{field['name']}]: {value}")
                    self.doc.Fields(field["name"]).ToggleSelect(value)
                    time.sleep(0.2)
        logger.debug("Exporting object...")
        chart.Export(Path(self.qlik_export_path).absolute(), ",")
        logger.debug(f"Object exported to {self.qlik_export_path}")
        return self.qlik_export_path

    def shutdown(self):
        try:
            if self._doc is not None:
                self._doc.CloseDoc()
                time.sleep(3)
                Exporter.doc_num -= 1
            if Exporter.doc_num <= 0 and Exporter._app is not None:
                Exporter._app.Quit()
                time.sleep(3)
        finally:
            pass

    def export(self):
        logger.info(f"Starting export {self.qlik_object_id} from {self.qlik_path_file}...")
        try:
            return self._runner()
        finally:
            self.shutdown()


if __name__ == "__main__":
    params = {
        "qlik_path_file": "C:\Program Files\QlikView\Examples\Documents\Qlik DataMarket.qvw",
        "qlik_export_path": "2012 Australia and Canada Sales.csv",
        "qlik_object_id": "CH194",
        "qlik_bookmark": "2012 Sales",
        "qlik_fields": [
            {
                "name": "Country",
                "values": ["Australia", "Canada"],
            }
        ]
    }
    Exporter(**params).export()
