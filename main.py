from docx import Document
from pprint import pprint 
import pyautogui
import clipboard
from time import sleep
answers = Document("answers.docx")

final_text = []

for paragraphs in answers.paragraphs:
    if paragraphs.text not in [""," ","\n"]:
        while "\xa0" in paragraphs.text:
            paragraphs.text = paragraphs.text.replace("\xa0"," ")
        final_text.append(paragraphs.text)

print(len(final_text))

counter = 1
one_text= ""
letter = 0
letter_list = [chr(i) for i in range(97,123)]
a = 0
texts = []
for paragraph in final_text:
    runs = len(paragraph)//148
    # print(runs)
    for i in range(runs): 
        one_text = str(counter)+"("+letter_list[letter]+")"+paragraph[a:a+146]
        texts.append(one_text)
        # print(a+5,a+150)
        a+=145
        letter +=1 
    a=0  
    letter = 0
    counter +=1 

print(len(texts))

sleep(5)
# for i in texts:
#     clipboard.copy(i)
#     sleep(0.1)
#     pyautogui.hotkey('ctrl','v')
#     pyautogui.press('enter')
#     sleep(0.1)
    # print(i)
    