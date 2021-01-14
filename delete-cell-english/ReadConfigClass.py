import configparser
import re


class ReadConfigClass():
    def __init__(self, path):
        self._config = configparser.ConfigParser()
        self._config.read(path, encoding='utf-8')

    def get_work_path(self):
        return self._config.get('Path', 'WorkPath', fallback='undefined Path')

    def get_sheet_name(self):
        sheet_names = []
        for sheet_name in self._config['SheetName']:
            sheet_names.append(self._config['SheetName'][sheet_name])
        return sheet_names

    def get_history(self):
        dict = {}
        cn = re.compile('([\u4e00-\u9fa5])')
        for values in self._config['History']:
            history = self._config['History'][values]
            index = history.find(cn.findall(history)[0])
            dict[history[0:index]] = history[index:]
        return dict

    def get_operation_key(self):
        return self._config.getint('Operation', 'operation', fallback=1)
