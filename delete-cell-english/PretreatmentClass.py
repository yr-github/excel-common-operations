import glob


# 当前只是读取所有excel传递
class PretreatmentClass():
    def __init__(self, path):
        self.path = path
        return

    def _list_all_excel(self):
        files_xls = glob.glob(self.path + "/*.xls*")
        return files_xls

    def pretreatment(self):
        files_xls = self._list_all_excel()
        return files_xls
