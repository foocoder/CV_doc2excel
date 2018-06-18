import re
import codecs
import csv

def parseDoc(file_name):
    doc_file = codecs.open(file_name, encoding="UTF-8")
    unicode_str = doc_file.read()
    ascii_str = unicode_str.encode('gbk','ignore')

    # Read name
    info_list = [i for i in re.findall("\s  mso-bidi-font-family:Courier;color:windowtext'>(.*)</span>",ascii_str)]
    phone_number = [i for i in re.findall("  color:windowtext'>(.*)<o:p></o:p></span>\s",ascii_str)]
    study_time = [i for i in re.findall("  </span>(\d{4}/\d{2}¨C\d{4}/\d{2})</span>",ascii_str)]
    print study_time
    info_list.insert(2, phone_number[0])
    # List infos
    #for i in info_list:
    #    print i

    return info_list

def traverseDir():
    import os

    doc_files_name = os.listdir(".\\doc")

    # TraverseDir
    info_mat = []
    for i in doc_files_name:
        info_list = parseDoc('.\\doc\\'+i)
        info_list.insert(0, i)
        info_mat.append(info_list)
        print i

    # Write csv
    with open(".\\excel\\a.csv","wb") as write_file:
        csv_writer = csv.writer(write_file)
        csv_writer.writerows(info_mat)


if __name__ == "__main__":
    traverseDir()
    
