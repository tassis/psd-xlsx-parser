import os
from psd_tools import PSDImage

class PSDParser:
    """
    PSDParser
    * 處理 PSD 相關的解析、輸出工作。
    """

    def __init__(self):
        self.file = None
        self.file_path = ''
        self.filename = ''

    def load(self, psd_path):
        """
        變更解析的 PSD 檔案。
        """
        path = os.path.join(psd_path[0], psd_path[1])
        self.filename = psd_path[1].split('.')[0]
        self.file = PSDImage.load(path)
        

    def get_layer_data_list(self):
        """
        遍歷 PSD 檔案 返回所有圖層名稱。
        """
        result = []
        for layer in reversed(self.file.layers):
            if 'ignore' in layer.name:
                continue
            result.append(self.__get_layer_data(layer))
        return result
    
    def export_png(self, layer_dict, pack_dir):
        """
        輸出圖片的功能。
        @param layer_data 需要輸出的圖層資料
        """
        out_path = f'{pack_dir}/{self.filename}'

        if not os.path.exists(out_path):
            os.makedirs(out_path)

        for layer in self.file.layers:
            if 'ignore' in layer.name:
                continue
            name = layer.name 
            if name in layer_dict:
                if layer_dict[name]['eng_name']:    
                    name = layer_dict[name]['eng_name']
            layer.as_PIL().save(f'{pack_dir}/{self.filename}/{name}.png') 

    def __get_layer_data(self, layer):
        """
        回傳 layer 相關資料
        """
        return {
            'tw_name': layer.name,
            'eng_name': '',
            'x': 720 - layer.bbox.x1,
            'y': 1280 - layer.bbox.y1
        }
    
    