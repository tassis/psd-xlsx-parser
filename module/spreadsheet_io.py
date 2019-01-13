import openpyxl
from templates.article import ArticleWorkbook, ArticleSheet

class SpreadsheetIO:
    """
    試算表的操作封裝
    """

    def __init__(self):
       self.file = None
       self.row_num = 0

    def load(self, xls_path):
        """
        讀取 Xlsx 檔案
        
        @param xls_path 試算表檔案位置
        """
        self.file = ArticleWorkbook(xls_path)

    def get_data(self):
        """
        取得 excel 表中所紀錄的資料
        """
        result = {}
        for item in self.file.data:
            key = item.tw_name
            result[key] = {
                'eng_name': item.eng_name,
                'x': item.x,
                'y': item.y
            }
        return result

    def export_table(self, data_list, path, header='', desc=''):
        """
        透過樣板輸出資料表
        """
        book = ArticleWorkbook()
        # 由於內容要求為 tuple ，因此轉換後再寫入。
        data_list = self.__dict_to_tuple(data_list)
        book.data.write(data_list, header, desc)
        book.save(path) 

    def compare_data(self, data_list):
        """
        與現有資料進行比對，若有誤差則以傳入的資料為主。
        """
        xls_data = self.get_data()
        for item in data_list:
            if item['tw_name'] in xls_data:
                key = item['tw_name']
                item['eng_name'] = xls_data[key]['eng_name']
        return data_list

    def __dict_to_tuple(self, data_list):
        """
        將內容轉換為元組
        """
        result = []
        for item in data_list:
            result.append(
                (self.row_num, item['tw_name'], item['eng_name'], item['x'], item['y'])
            )
            self.row_num = self.row_num + 1
        return result

    def list_to_dict(self, data_list):
        result = {}
        print(data_list)
        for item in data_list:
            key = item['tw_name']
            result[key] = {
                'eng_name': item['eng_name'],
                'x': item['x'],
                'y': item['y']
            }
        return result