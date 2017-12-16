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
    work_folder = 'C:\\Users\\atsuk\\PycharmProjects\\study\\utility\\work'
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
        point = [sheet['D22'].value, sheet['D23'].value, sheet['D24'].value]
        midashi_list = ['ア', 'イ', 'ウ', 'エ', 'オ', 'カ', 'キ', 'ク', 'ケ', 'コ']
        midashi_position_list = ['D27', 'D30', 'D34', 'D37', 'D41', 'D44', 'D49', 'D53', 'D57', 'D61']

        for i in range(len(midashi_list)):
            if point[0] == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E22'].value:
                    print(midashi_position_list[i] + ':見出しがポイント1と同一です')
            if point[1] == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E23'].value:
                    print(midashi_position_list[i] + ':見出しがポイント2と同一です')
            if point[2] == midashi_list[i]:
                if sheet[midashi_position_list[i]].value == sheet['E24'].value:
                    print(midashi_position_list[i] + ':見出しがポイント3と同一です')

        # ⑦
        # 「自社」「弊社」→「同社」に統一
        # 取引先「様」トル
        # 「です」「ます」→「である」に統一
        # 「明治」「大正」「昭和」「平成」→西暦に統一
        # 掲載するところだけ選出
        text_position_list = ['C21', 'E22', 'E23', 'E24', 'C25']
        midashi_num_list = []
        naiyo_num_list = []

        for j in point:
            if j == 'ア':
                text_position_list.append('D27')
                text_position_list.append('D28')
                midashi_num_list.append('D27')
                naiyo_num_list.append('D28')
            if j == 'イ':
                text_position_list.append('D30')
                text_position_list.append('D31')
                midashi_num_list.append('D30')
                naiyo_num_list.append('D31')
            if j == 'ウ':
                text_position_list.append('D34')
                text_position_list.append('D35')
                midashi_num_list.append('D34')
                naiyo_num_list.append('D35')
            if j == 'エ':
                text_position_list.append('D37')
                text_position_list.append('D38')
                midashi_num_list.append('D37')
                naiyo_num_list.append('D38')
            if j == 'オ':
                text_position_list.append('D41')
                text_position_list.append('D42')
                midashi_num_list.append('D41')
                naiyo_num_list.append('D42')
            if j == 'カ':
                text_position_list.append('D44')
                text_position_list.append('D45')
                midashi_num_list.append('D44')
                naiyo_num_list.append('D45')
            if j == 'キ':
                text_position_list.append('D49')
                text_position_list.append('D50')
                midashi_num_list.append('D49')
                naiyo_num_list.append('D50')
            if j == 'ク':
                text_position_list.append('D53')
                text_position_list.append('D54')
                midashi_num_list.append('D53')
                naiyo_num_list.append('D54')
            if j == 'ケ':
                text_position_list.append('D57')
                text_position_list.append('D58')
                midashi_num_list.append('D57')
                naiyo_num_list.append('D58')
            if j == 'コ':
                text_position_list.append('D61')
                text_position_list.append('D62')
                midashi_num_list.append('D61')
                naiyo_num_list.append('D62')

        ng_word = ['自社', '弊社', '様', 'です', 'ます', '明治', '大正', '昭和', '平成'
                    , '１０', '２０', '３０', '４０', '５０', '６０', '７０', '８０', '９０', '００'
                    , '１１', '２１', '３１', '４１', '５１', '６１', '７１', '８１', '９１', '０１'
                    , '１２', '２２', '３２', '４２', '５２', '６２', '７２', '８２', '９２', '０２'
                    , '１３', '２３', '３３', '４３', '５３', '６３', '７３', '８３', '９３', '０３'
                    , '１４', '２４', '３４', '４４', '５４', '６４', '７４', '８４', '９４', '０４'
                    , '１５', '２５', '３５', '４５', '５５', '６５', '７５', '８５', '９５', '０５'
                    , '１６', '２６', '３６', '４６', '５６', '６６', '７６', '８６', '９６', '０６'
                    , '１７', '２７', '３７', '４７', '５７', '６７', '７７', '８７', '９７', '０７'
                    , '１８', '２８', '３８', '４８', '５８', '６８', '７８', '８８', '９８', '０８'
                    , '１９', '２９', '３９', '４９', '５９', '６９', '７９', '８９', '９９', '０９']

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
                        text_val = text_val[ng_find + 2:]
                    else:
                        check_flg = False


        # 字数チェック
        # キャッチフレーズ 30字以上57字以下
        if len(sheet['C21'].value) < 30:
            print('C21:キャッチフレーズの字数が少なすぎます')
        elif len(sheet['C21'].value) > 57:
            print('C21:キャッチフレーズの字数が多すぎます')

        # 取り組みの要約3項目 30字以上48字以下
        if len(sheet['E22'].value) < 30:
            print('E22:取組の要約1の字数が少なすぎます')
        elif len(sheet['E22'].value) > 48:
            print('E22:取組の要約1の字数が多すぎます')
        if len(sheet['E23'].value) < 30:
            print('E23:取組の要約2の字数が少なすぎます')
        elif len(sheet['E23'].value) > 48:
            print('E23:取組の要約2の字数が多すぎます')
        if len(sheet['E24'].value) < 30:
            print('E24:取組の要約3の字数が少なすぎます')
        elif len(sheet['E24'].value) > 48:
            print('E24:取組の要約3の字数が多すぎます')

        # 会社概要 160字以上228字以下
        if len(sheet['C25'].value) < 160:
            print('C25:会社概要の字数が少なすぎます')
        elif len(sheet['C25'].value) > 228:
            print('C25:会社概要の字数が多すぎます')

        # 見出し 12字以上28字以下
        for i in midashi_num_list:
            if len(sheet[i].value) < 12:
                print(i + ':見出しの字数が少なすぎます')
            elif len(sheet[i].value) > 28:
                print(i + ':見出しの字数が多すぎます')

        # 本文 180字以上228字以下
        for i in naiyo_num_list:
            if len(sheet[i].value) < 180:
                print(i + ':本文の字数が少なすぎます')
            elif len(sheet[i].value) > 228:
                print(i + ':本文の字数が多すぎます')