import random
from utils import wxwx, Source_Dataset, read_txt, get_sort_ByProbability_OfOneProperty, txt_clearAll, xlsx_clearAll
import os

sep = os.sep

max_num_case = 60
Min_Max_LengthCharacter_str = "30_60"
Novel_Name = 'ThreeBody'

root_GPTWord = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords'

# path_txt_ori = root_GPTWord + '\data\Txt\Raw\Rickshaw Boy\LTXZ_1000_' + Min_Max_LengthCharacter_str + '.txt'
# path_txt_cut = root_GPTWord + '\data\Txt\Preprocess\Rickshaw Boy\LTXZ_1000_' + Min_Max_LengthCharacter_str + '.txt.txt'
path_txt_ori = root_GPTWord + '\data\Txt\Raw\ThreeBody_10000_30_60.txt'
path_txt_cut = root_GPTWord + '\data\Txt\Preprocess\ThreeBody_10000_30_60.txt.txt'

path_dataset_xlsx = root_GPTWord + '\data\Root_Dataset_xlsx\dataset_xlsx_20230426_Main.xlsx'
Num_Multi_Words = 1  # 只能等于1，否则有Bug！！！


for type_prop_list in [['v']]:
    path_txt_GPT = root_GPTWord + '\data\Experiment' + sep + Novel_Name + sep + 'TxtGPT' + sep + Novel_Name + '_Dataset_Case_' + str(
        max_num_case) + '_' + ''.join(
        type_prop_list) + '_' + str(Num_Multi_Words) + '.txt'
    path_preprocess_xlsx = root_GPTWord + '\data\Experiment' + sep + Novel_Name + sep + 'Exp' + sep + Novel_Name + '_Dataset_Case_' + str(
        max_num_case) + '_' + ''.join(
        type_prop_list) + '_' + str(Num_Multi_Words) + '.xlsx'

    Words_Length = 2
    txt_GPT_ClearAll = True
    xlsx_result_ClearAll = True

    if txt_GPT_ClearAll:
        txt_clearAll(path_txt_GPT)
    if xlsx_result_ClearAll:
        xlsx_clearAll(path_preprocess_xlsx)

    ww = Source_Dataset(path_xlsx_dataset=path_dataset_xlsx)

    wx_result = wxwx(path_xlsx=path_preprocess_xlsx)
    wx_result.write_one_value("原句", 1, 1)

    for i in range(Num_Multi_Words):
        wx_result.write_one_value("原词" + str(i), 1, 1 + i * 4 + 1)
        wx_result.write_one_value("新词" + str(i), 1, 1 + i * 4 + 2)
        wx_result.write_one_value("原词词频" + str(i), 1, 1 + i * 4 + 3)
        wx_result.write_one_value("新词词频" + str(i), 1, 1 + i * 4 + 4)

    list_sentence_ori = read_txt(path_txt_ori).split('\n')
    list_sentence_cut = read_txt(path_txt_cut).split('\n')

    len_list_sentence_ori = len(list_sentence_ori)
    len_list_sentence_cut = len(list_sentence_cut)
    if len_list_sentence_ori != len_list_sentence_cut:
        print('Num Wrong! len_list_sentence_ori != len_list_sentence_cut !!!')

    iter_case_saved = 0
    GPT_txt = open(path_txt_GPT, 'a+', encoding='utf-8')

    exist_index = []
    exist_word = []
    for a in range(len(list_sentence_ori)):
        if iter_case_saved < max_num_case:
            help_index = True
            while help_index:
                the_a = random.randrange(len(list_sentence_ori))
                if the_a not in exist_index:
                    exist_index.append(the_a)
                    one_sentence_ori = list_sentence_ori[the_a]
                    one_sentence_cut = list_sentence_cut[the_a]
                    out = get_sort_ByProbability_OfOneProperty(sentence_OriCut=one_sentence_cut, ww_Source_Dataset=ww,
                                                               type_prop_list=type_prop_list)

                    WordSeries_change = out[0]
                    word_ori = WordSeries_change.word
                    if word_ori not in exist_word and len(word_ori) == Words_Length:
                        exist_word.append(word_ori)
                        this_row_value_list = [one_sentence_ori]
                        GPT_txt_content = '('
                        this_word = word_ori
                        this_row_value_list += [this_word, None, None, None]
                        GPT_txt_content += this_word

                        wx_result.append_value_list(this_row_value_list)

                        GPT_txt_content += ')' + ';' + '(' + one_sentence_ori + ')' + '\n'
                        GPT_txt.write(GPT_txt_content)
                        print('iter_case_saved=',iter_case_saved)

                        iter_case_saved += 1
                        help_index = False
                        break

        else:
            break

    wx_result.save_xlsx_cover(wx_result.path_xlsx)

    GPT_txt.close()
