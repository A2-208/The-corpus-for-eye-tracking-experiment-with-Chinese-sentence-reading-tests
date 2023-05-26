import scipy.stats
from scipy.stats import ttest_rel
from utils import wxwx
import numpy as np

for type_type in [['n', 'v', 'a'], ['n'], ['v'], ['a']]:
    root_ChatGPT_document = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords'
    path_result_xlsx_0 = root_ChatGPT_document + '\Paper\Dataset_Show\Source\LTXZ_Dataset_Out_' + "".join(type_type) + '.xlsx'
    print('path_result_xlsx_0=',path_result_xlsx_0)
    ww_0 = wxwx(path_xlsx=path_result_xlsx_0)

    ori = []
    new = []
    for i in range(2, ww_0.sheets[0].max_row + 1):
        one_ori_prob = ww_0.get_one_value(i, 4)
        one_new_prob = ww_0.get_one_value(i, 5)

        ori.append(one_ori_prob)
        new.append(one_new_prob)

    T_Test_t, T_Test_p = ttest_rel(ori, new)
    Wilcoxon = scipy.stats.ranksums(ori, new)
    ori_arr = np.array(ori)
    new_arr = np.array(new)

    print('type_type=',type_type)
    print('Aver Ori=', round(float(np.mean(ori_arr)), 3),'  Aver New=', round(float(np.mean(new_arr)), 3))
    print('Std Ori=', round(float(np.std(ori_arr)), 3), '  Std New=', round(float(np.std(new_arr)), 3))
    print('Wilcoxon-p-value=', Wilcoxon[1], '  T_Test-p-value=', T_Test_p)
