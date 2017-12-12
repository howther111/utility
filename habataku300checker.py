import openpyxl
import os  # osモジュールのインポート

def year_change(gengo, year):
    if gengo == '明治':
        if year != '元' and year != '':
            ans = str(int(year) + 1867)
            return ans
        else:
            return '1988'
    elif gengo == '大正':
        if year != '元' and year != '':
            ans = str(int(year) + 1911)
            return ans
        else:
            return '1912'
    elif gengo == '昭和':
        if year != '元' and year != '':
            ans = str(int(year) + 1925)
            return ans
        else:
            return '1926'
    elif gengo == '平成':
        if year != '元' and year != '':
            ans = str(int(year) + 1988)
            return ans
        else:
            return '1989'

if __name__ == '__main__':
    # os.listdir('パス')
    # 指定したパス内の全てのファイルとディレクトリを要素とするリストを返す
    work_folder = 'C:\\Users\\atsuk\\PycharmProjects\\study\\ExcelEditor\\work'
    files = os.listdir(work_folder)
    for file_name in files:
        wb = openpyxl.load_workbook(work_folder + '\\' + file_name)
        sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])
        print(file_name + ' チェック開始')

        # ①社名と株式会社の間は詰める
        company_name = sheet['C11'].value
        if '株式会社 ' in company_name:
            print('C11:社名と株式会社の間は詰める')
        elif '株式会社　' in company_name:
            print('C11:社名と株式会社の間は詰める')
        elif ' 株式会社' in company_name:
            print('C11:社名と株式会社の間は詰める')
        elif '　株式会社' in company_name:
            print('C11:社名と株式会社の間は詰める')
        if '株式会社' not in company_name:
            print('C11:「株式会社」の表記なし')

        # ②TEL、FAX番号にはハイフン入れる
        tel = sheet['C15'].value
        if '-' not in tel:
            print('C15:電話番号にハイフン入れる')

        fax = sheet['I15'].value
        if '-' not in fax:
            print('I15:FAX番号にハイフン入れる')

        url = sheet['C16'].value
        if 'http' not in url:
            print('C16:URL記載なし')

        # ③代表者の姓と名の間にスペース入れる
        name = sheet['C13'].value
        if '　' not in name:
            print('C13:代表者の姓と名の間に全角スペース入れる')

        # ④資本金の表記「,」入れる（正常に動作せず）
        capital = sheet['K18'].value
        if len(str(capital)) > 3 and ',' not in str(capital):
            print('K18:資本金の表記「,」入れる ' + str(capital))

        # ⑤判別不能
        # ⑥各3項目の見出しと同一文章の場合は修正を依頼
        point_1 = sheet['D22'].value
        point_2 = sheet['D23'].value
        point_3 = sheet['D24'].value
        midashi_list = ['ア', 'イ', 'ウ', 'エ', 'オ', 'カ', 'キ', 'ク', 'ケ', 'コ']
        midashi_position_list = ['D27', 'D30', 'D34', 'D37', 'D41', 'D44', 'D49', 'D53', 'D57', 'D61']

        for i in range(10):
            if point_1 == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E22'].value:
                    print(midashi_position_list[i] + ':見出しがポイント1と同一です')
            if point_2 == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E23'].value:
                    print(midashi_position_list[i] + ':見出しがポイント2と同一です')
            if point_3 == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E24'].value:
                    print(midashi_position_list[i] + ':見出しがポイント3と同一です')

        # ⑦
        # 「自社」「弊社」→「同社」に統一
        # 取引先「様」トル
        # 「です」「ます」→「である」に統一
        # 「明治」「大正」「昭和」「平成」→西暦に統一

        text_position_list = ['E22', 'E23', 'E24', 'C25', 'D27', 'D28', 'D29', 'D30', 'D31', 'D32'
            , 'D34', 'D35', 'D36', 'D37', 'D38', 'D39', 'D41', 'D42', 'D43', 'D44', 'D45', 'D46'
            , 'D49', 'D50', 'D51', 'D53', 'D54', 'D55', 'D57', 'D58', 'D59', 'D61', 'D62', 'D63']

        ng_word = ['自社', '弊社', '様', 'です', 'ます', '明治', '大正', '昭和', '平成']

        for text in text_position_list:
            for ng in ng_word:
                ng_find = str(sheet[text].value).find(ng)

                text_val = str(sheet[text].value)
                check_flg = True
                while check_flg:
                    ng_find = text_val.find(ng)
                    if ng_find != -1:
                        front = 0
                        if (ng_find > 5):
                            front = 5
                        print(text + ':「' + ng + '」あり ' + text_val[ng_find - front: ng_find + 10])
                        # 元号を西暦に変換
                        if ng == '明治' or ng == '大正' or ng == '昭和' or ng == '平成':
                            if text_val[ng_find + 3] == '年':
                                year_pos = ng_find + 3
                                gengo_year = text_val[ng_find + 2:year_pos]
                                seireki_year = year_change(ng, gengo_year)
                                print(ng + gengo_year + '年→' + seireki_year + '年')
                            elif text_val[ng_find + 4] == '年':
                                year_pos = ng_find + 4
                                gengo_year = text_val[ng_find + 2:year_pos]
                                seireki_year = year_change(ng, gengo_year)
                                print(ng + gengo_year + '年→' + seireki_year + '年')
                            elif text_val[ng_find + 5] == '年':
                                year_pos = ng_find + 5
                                gengo_year = text_val[ng_find + 2:year_pos]
                                seireki_year = year_change(ng, gengo_year)
                                print(ng + gengo_year + '年→' + seireki_year + '年')
                        text_val = text_val[ng_find + 1:]
                    else:
                        check_flg = False


        # 字数チェック 後で実装