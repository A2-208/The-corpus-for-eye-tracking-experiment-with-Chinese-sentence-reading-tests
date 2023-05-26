from utils import TransDownload_To_DatasetXlsx

'''''
 The root_download contains the data(.txt) downloaded from 'http://corpus.zhonghuayuwen.org'
 You need to do it in your way to acquire the ones with frequency < 50 from 'http://corpus.zhonghuayuwen.org'
 And the "selenium" is helpful, and I believe that you can make it, which is easy. 
 '''''

path_xlsxLoad = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords\data\root_dataset_xlsx\dataset_xlsx_20230402.xlsx'
root_download = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords\data\DownloadDataRoot\20230401_Total5000'

do_Before_checkDataset = False
do_After_checkDataset = True

TransDownload_To_DatasetXlsx(path_xlsxLoad=path_xlsxLoad,
                             root_download=root_download,
                             do_Before_checkDataset=do_Before_checkDataset,
                             do_After_checkDataset=do_After_checkDataset,
                             )
