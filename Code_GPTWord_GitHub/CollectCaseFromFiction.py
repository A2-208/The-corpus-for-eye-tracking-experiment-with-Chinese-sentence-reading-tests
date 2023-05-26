from utils import read_txt, txt_clearAll
import random

Num_case = 1000
min_len = 30
max_len = 60

root_GPTWord = r'D:\Work_Skyer\Lab\Doctor_ZYJ\GPTWords'

path_SourceTxt = root_GPTWord + '\Reference\ContentSource\LTXZ.txt'
path_txt_ori = root_GPTWord + '\data\Txt\Raw\Rickshaw Boy\LTXZ_' + str(Num_case) + '_' + str(min_len) + '_' + str(max_len) + '.txt'

str_SourceTxt = read_txt(path_SourceTxt)
special_str = '----------'
list_S = str_SourceTxt.split('\n')

BannedList = [list_S[0], list_S[1], list_S[2], '', ' ', '  ', '   ']
New_list_S = []
for one_s in list_S:
    if one_s in BannedList:
        pass
    elif len(one_s) < 3:
        pass
    else:
        one_s = one_s[3:]
        New_list_S.append(one_s)

New_Str_S = ''.join(New_list_S)
F_list = ['，', '。']
banned_character_list = ['①', '②', '③', 'BaptisteRenéRobinet，1735—1820）', 'BaptisteRenéRobinet，1735—1820）']

the_list = New_Str_S.split('。')

the_list_Temp = []
for i in the_list:
    one_i = []
    for one_char in list(i):
        if one_char not in banned_character_list:
            one_i.append(one_char)
    one_i = "".join(one_i)
    the_list_Temp.append(one_i)

the_list = the_list_Temp
exist_list = []

txt_clearAll(path_txt_ori)
txt_Raw = open(path_txt_ori, 'a+', encoding='utf-8')
for i in range(Num_case):
    help_index = True
    while help_index:
        print(len(exist_list))

        index_r = random.randrange(0, len(the_list))
        one_sentence = the_list[index_r]

        one_sentence_list = []
        the_phrase_0 = one_sentence.split('？')
        for a in the_phrase_0:
            the_phrase_1 = one_sentence.split('！')
            for b in the_phrase_1:
                the_phrase_2 = one_sentence.split('。')
                for c in the_phrase_2:
                    the_phrase_3 = one_sentence.split('，')
                    for d in the_phrase_3:
                        one_sentence_list.append(d)

        random.shuffle(one_sentence_list)

        for one_sss in one_sentence_list:
            len_one_sentence = len(one_sentence)
            if max_len > len_one_sentence > min_len and (
                    special_str not in one_sentence) and one_sentence not in exist_list:
                txt_Raw.write(one_sentence)
                exist_list.append(one_sentence)
                if i == Num_case - 1:
                    pass
                else:
                    txt_Raw.write('\n')
                help_index = False
            break

txt_Raw.close()
