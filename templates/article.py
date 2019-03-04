from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet
from openpyxl_templates.table_sheet.columns import CharColumn, IntColumn

class ArticleSheet(TableSheet):
    primekey = IntColumn("編號")
    tw_name = CharColumn("圖層名稱")
    eng_name = CharColumn("檔案名稱")
    x = CharColumn()
    y = CharColumn()
    width = CharColumn('寬')
    height = CharColumn('高')

class ArticleWorkbook(TemplatedWorkbook):
    data = ArticleSheet("data")
    