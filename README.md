準備環境 

1.安裝 python 3.6.6 以上的版本  
2.使用指令 pip install -r requirement.txt  安裝所有依賴套件  
3.將 PSD 拖曳入資料夾  
4.運行 setup 腳本  
  
運行說明  

在沒有 xlsx 檔案存在的情況下運行時，第一次會先將 psd 檔案中的圖層資訊輸出並儲存成與 psd 相同名稱的 xlsx 檔案之後中斷，待使用者確認後再次運行才會輸出 png 檔案於 __pack__ 資料夾中。

問與答

Q: 我點了兩下後只輸出了一個 xlsx 檔案卻沒有 __pack__。
A: 在 xlsx 檔案中填入需要更改的檔案名稱(不填寫的話預設會以圖層名稱為主)後，在次運行腳本即可。

Q: 視窗提示 No module XXX 怎麼辦？
A: 請確認已經安裝了 Python 3.6 以上的版本並於資料夾內運行過 pip install -r requipment.txt。