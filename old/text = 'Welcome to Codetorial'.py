text = 'Welcome to Codetorial'

#pos_e_last = text.rfind('!')
if text.rfind('~') == -1:
    print("1111" , text.rfind('!'))
else :
    print("2222" , text.rfind('!'))




pos_e_first = text.find('e')
if pos_e_first == -1:
    print("3333" , pos_e_first)
else :
    print("4444" , pos_e_first)



pos_to_last = text.rfind('to')
print(pos_to_last)

pos_to_first = text.find('to')
print(pos_to_first)

