# -*- coding: UTF-8 -*-
import csv, sys, optparse
from openpyxl import load_workbook

# GENIAN NAC 관리 -> 노드 -> 전체노드 -> 내보내기 선택
# python excel_to_cvs2 '내보낸 파일'

class LookupTableCreate:

    def Convert_Table(self, outputFile, worksheet):
        with open(outputFile, 'w', newline='') as user_file:  # newline='' 입력 시 저장되는 파일에 공백이 제거됨
            userlists = []
            saveFile = csv.writer(user_file, quoting=csv.QUOTE_NONNUMERIC)  # quoting=csv.QUOTE_NONNUMERIC "" 선언
            for row in worksheet:
                if row[11].value in ['', None, 'None']: #사용자 이름 없는 것은 건너뜀
                    pass
                else:
                    if [row[5].value, row[11].value] not in userlists: # 중복 제거
                        userlists.append([row[5].value, row[11].value]) # 5=IP, 11=사용자이름

            for csv_row in userlists: # 엑셀로 저장
                saveFile.writerow(csv_row)

        return print("[!] 변환완료")


    def Make_cvs(self, CvsFile):
        inputFile = CvsFile
        outputFile = 'gabiauser.csv'  # 저장되는 파일이름
        workbook = load_workbook(filename=inputFile, data_only=True)  # data_only = 함수 실행 결과 값 추출
        worksheet = workbook['노드 관리']  # 읽는 시트
        self.Convert_Table(outputFile, worksheet)


def main():
    print("[+] Genian Node To Graylog LookupTable cvs")
    usage = "[+] Usage : Python %Prog [-c Convert File]"
    parser = optparse.OptionParser(usage=usage)
    parser.add_option('-c', '--create', dest='CvsFile', help='specify a compare file')

    (options, args) = parser.parse_args()
    if not options.CvsFile:
         print(parser.usage)
         sys.exit(0)

    Create = LookupTableCreate()
    Create.Make_cvs(options.CvsFile)
    sys.exit(0)


if __name__ == "__main__":
    main()