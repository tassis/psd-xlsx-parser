# coding=utf-8
import os
import sys
import logging
import shutil
from module.psd_parser import PSDParser
from module.spreadsheet_io import SpreadsheetIO
# initial logger system.


# initial parameter.
parser = PSDParser()
sheetIO = SpreadsheetIO()

def get_psd_list(dir_path='.'):
    """
    掃描資料夾內所有的 psd 檔案
    @params dir_path 資料夾路徑
    """
    item_list = []
    for item in os.listdir(dir_path):
        if '.psd' in item:
            item_list.append([dir_path, item])
    return item_list

def get_xlsx_data(psd_path):
    """
    檢查 psd 檔案是否有對應的 xls 檔案
    
    @params psd_path 檔案路徑
    """
    # 若不是psd檔案則直接跳過
    if '.psd' not in psd_path[1]:
        return
    # 取得檔名並轉換成excel檔名
    basename = '{0}.{1}'.format(psd_path[1].split('.')[0], 'xlsx')
    excel_path = os.path.join(psd_path[0], basename)
    # 讀取 PSD 的圖層資料
    parser.load(psd_path)
    layer_list = parser.get_layer_data_list()
    # 檢查 xlsx 是否存在，不存在則生成，若存在則進行資料比對。
    if not os.path.isfile(excel_path):
        sheetIO.export_table(layer_list, excel_path, basename, psd_path[1])
        return None
    # xlsx 存在刷新內容並存檔
    sheetIO.load(excel_path)
    # 比對資料，以 psd 內的資料為主。
    sheet_data = sheetIO.compare_data(layer_list)
    sheetIO.export_table(sheet_data, excel_path, basename, psd_path[1])
    print(f'load xlsx {excel_path}')
    return sheetIO.get_data()

def export_png(layer_list, dir_path):
    """
    輸出圖片的功能。

    @param layer_list 輸出時使用的比對資料
    """
    pack_dir = os.path.join(dir_path, '__pack__')
    if os.path.exists(pack_dir):
        shutil.rmtree(pack_dir)
    
    if not os.path.exists(pack_dir):
        os.makedirs(pack_dir)

    # layer_dict = sheetIO.list_to_dict(layer_list)
    parser.export_png(layer_list, pack_dir)

def main(path):
    """
    運行的主要流程
    """
    if os.path.isfile(path):
        # 為單一檔案的情況，先將路徑分割再解析。
        dirname = os.path.dirname(path) 
        basename = os.path.basename(path)
        xlsx_data = get_xlsx_data([dirname, basename])
        export_png(xlsx_data, dirname)

    elif os.path.isdir(path):
        # 為資料夾的情況
        psd_list = get_psd_list(path)
        # 遍歷所有 psd 檔案並檢索對應 xlsx 。
        for item in psd_list:
            xlsx_data = get_xlsx_data(item)
            export_png(xlsx_data, path)

if __name__ == '__main__':
    if (len(sys.argv) > 1):
        # 有路徑參數的情況
        main(sys.argv[1])
    else:
        # 沒有路徑參數的情況
        main('.')