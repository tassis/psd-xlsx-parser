from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet
from openpyxl_templates.table_sheet.columns import CharColumn, IntColumn

class ArticleSheet(TableSheet):
    primekey = IntColumn("編號")
    tw_name = CharColumn("中文名稱")
    eng_name = CharColumn("英文名稱")
    x = CharColumn()
    y = CharColumn()

class ArticleWorkbook(TemplatedWorkbook):
    data = ArticleSheet("data")
    