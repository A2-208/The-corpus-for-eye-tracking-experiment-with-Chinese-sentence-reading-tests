from utils import wxwx

path_temp_xlsx = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords\Paper\Dataset_Show\Temp.xlsx'
path_result_dataset_xlsx = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords\Paper\Dataset_Show\Rickshaw Boy_Corpus_Dataset_V2.xlsx'

wx_result_dataset_xlsx = wxwx(path_xlsx=path_temp_xlsx)
wx_result_dataset_xlsx.append_value_list(['High Frequency Word','Low Frequency Word','Word Frequency ‰ (High)','Word Frequency ‰ (Low)','Sentence (High Frequency)','Sentence (Low Frequency)'])

name_list = ['n', 'v', 'a', 'nva']
help_index = 1
exist_ori_sentence = []

for one_name in name_list:
    path_xlsx_source = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords\Paper\Dataset_Show\Source\LTXZ_Dataset_Out_' + one_name + '.xlsx'
    xx_source = wxwx(path_xlsx=path_xlsx_source)
    for row_index in range(2, 92):
        sentence = xx_source.get_one_value(row_index=row_index, column_index=1)
        if sentence not in exist_ori_sentence:
            ori_word = xx_source.get_one_value(row_index=row_index, column_index=2)
            new_word = xx_source.get_one_value(row_index=row_index, column_index=3)
            ori_word_f = xx_source.get_one_value(row_index=row_index, column_index=4)
            new_word_f = xx_source.get_one_value(row_index=row_index, column_index=5)

            if ori_word == new_word:
                print(ori_word, '  ', one_name, '  ', row_index)

            if float(ori_word_f) > float(new_word_f):
                # print('sentence=',sentence)
                # print('ori_word=',ori_word)

                ori_word_start_index = sentence.index(ori_word)
                old_word_index_list = [ori_word_start_index, ori_word_start_index + len(ori_word)]
                part_before = sentence[:old_word_index_list[0]]
                part_after = sentence[old_word_index_list[1]:]
                new_sentence = ''.join([part_before, new_word, part_after])

                wx_result_dataset_xlsx.append_value_list([ori_word, new_word, ori_word_f, new_word_f, sentence, new_sentence])
                exist_ori_sentence.append(sentence)
                # print(help_index)
                help_index += 1

# wx_result_dataset_xlsx.save_xlsx_cover(path_save=path_result_dataset_xlsx)
