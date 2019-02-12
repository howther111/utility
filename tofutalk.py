# coding: utf-8
text_array = []

filename = input("ファイル名を入力してください:")

f = open(filename + ".txt", encoding='utf-8')
line = f.readline()
text_array.append(line)
while line:
    line = f.readline()
    if line != "":
        text_array.append(line)
f.close()

count = 0

text = ""

for i in text_array:
    try:
        count = count + 1
        serifu = i.split("：")
        dicebotflg = False
        if len(serifu) < 2:
            serifu = serifu[0].split(" : ")
            dicebotflg = True

        jnum = 0
        for j in serifu:
            if jnum > 1:
                serifu[1] = serifu[1] + "：" + j
            jnum = jnum + 1

        if "「" in serifu[1]:
            serifu[1] = serifu[1].replace("\n", "")
        else:
            serifu[0] = serifu[0] + "Ｎ"
            serifu[1] = "「" + serifu[1].replace("\n", "") + "」"

        if dicebotflg:
            serifu[0] = "ダイスボット"

        text = text + serifu[0] + serifu[1] + "\n"

        print(serifu[0] + serifu[1])

    except:
        print(serifu)

f = open('output_' + filename + '.txt', encoding='utf-8', mode='w')
f.write(text)
f.close()
