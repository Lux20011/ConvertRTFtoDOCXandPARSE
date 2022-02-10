import datetime
import os.path
import pathlib
import random
import re
import docx2txt
import win32com.client


baseDir = '.' # Starting directory for directory walk
some_dir = os.listdir('.')
word = win32com.client.Dispatch("Word.application")

for dir_path, dirs, files in os.walk(baseDir):
    for file_name in files:

        file_path = os.path.join(dir_path, file_name)
        file_name, file_extension = os.path.splitext(file_path)

        if "~$" not in file_name:
            if file_extension.lower() == '.rtf': #
                docx_file = '{0}{1}'.format(file_name, '.docx')

                if not os.path.isfile(docx_file): # Skip conversion where docx file already exists

                    file_path = os.path.abspath(file_path)
                    docx_file = os.path.abspath(docx_file)
                    try:
                        wordDoc = word.Documents.Open(file_path)
                        wordDoc.SaveAs2(docx_file, FileFormat = 16)
                        wordDoc.Close()
                        os.remove(file_path)
                    except Exception as e:
                        print('Failed to Convert: {0}'.format(file_path))
                        print(e)

input('PRESS TO CONTINUE')


for newdocxfiles in some_dir:
    counter = 0
    file_extansion = pathlib.Path(newdocxfiles).suffix
    if file_extansion in ('.docx','.doc'):

        my_text = docx2txt.process(newdocxfiles)
        h_data = re.findall(r'(?:[\d]{2}(?!\d{2}).[\d]{2}.[\d]{4}){1}',my_text)
        h_head = re.findall(r'(ВЫПИСКА|Ведомость|ПРИЛОЖЕНИЕ К ВЫПИСКЕ|ПРИЛОЖЕНИЕ К ВЫПИСКE)',my_text,flags=re.IGNORECASE)
        h_head_descr = re.findall(r'(кассовых поступлений в бюджет|из лицевого счета бюджета|из лицевого счета главного распорядителя|по движению свободного остатка|из лицевого счета администратора доходов бюджета|из казначейского счета|из лицевого счета получателя бюджетных средств)',my_text,flags=re.IGNORECASE)
        h_name = re.findall(r'(Славянский район|Бюджет Краснодарского края|Ачуевского сельского поселения|Протокского сельского поселения|Прикубанского сельского поселения|Славянского городского поселения|Рисового сельского поселения|Целинного сельского поселения|Маевского сельского поселения|сельского поселения Голубая Нива)',my_text,flags=re.IGNORECASE)
        h_data_new = h_data[0]
        h_head_new = h_head[0]
        h_head_descr_new = h_head_descr[0]
        h_name_new = h_name[0]
        h_new_filename_end = (str(h_data_new)+" "+str(h_head_new)+' '+str(h_head_descr_new)+" "+str(h_name_new)+file_extansion)
        if h_new_filename_end == newdocxfiles:
            counter+=1
            print(str(counter))
            h_new_filename_end = (str(h_data_new) + " " + str(h_head_new) + ' ' + str(h_head_descr_new) + " " + str(h_name_new)+str(counter) + file_extansion)

        else:
            print("NO")
        print(str(h_new_filename_end))
        print(str(newdocxfiles))

        os.rename(newdocxfiles, h_new_filename_end)
        #print(str(h_data_new)+' '+str(h_head_new) +' '+str(h_head_new_descr)+' '+str(h_name)+' '+str(h))

input('PRESS ENTER TO END')
