import ReadConfigClass
import PretreatmentClass
import DeleteEnglishClass
import TranslationByHistoryClass
import os


if __name__ == '__main__':
    rd_conf_obj = ReadConfigClass.ReadConfigClass(os.getcwd()+"\\config.ini")
    pretreatment_obj = PretreatmentClass.PretreatmentClass(rd_conf_obj.get_work_path())
    files_xls = pretreatment_obj.pretreatment()
    if 1 == rd_conf_obj.get_operation_key():
        print("开始执行删除英文操作...")
        print(files_xls)
        delete_eng_obj = DeleteEnglishClass.DeleteEnglishClass(files_xls, rd_conf_obj.get_sheet_name())
        delete_eng_obj.run()
    if 2 == rd_conf_obj.get_operation_key():
        print("开始执行翻译英文操作...")
        print(files_xls)
        translation_obj = TranslationByHistoryClass.TranslationByHistoryClass(rd_conf_obj.get_history(),
                                                                              files_xls, rd_conf_obj.get_sheet_name())
        translation_obj.run()
