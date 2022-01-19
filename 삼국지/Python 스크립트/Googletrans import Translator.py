from itranslate import itranslate as itrans

itrans("test this and that")  # '测试这一点'

# new lines are preserved, tabs are not
itrans("test this \n\nand test that \t and so on")
# '测试这一点\n\n并测试这一点等等'

print(itrans("毋丘倹", from_lang="zh-Hans", to_lang="en"))  # 'Testen Sie das und das'
