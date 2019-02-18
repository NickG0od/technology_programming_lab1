'''
! Файл, для чтения текста из doc / docx файлов.
-----------------------------------------------
* Подсчёт кол-ва предложений в тексте.
* Подсчёт кол-ва слов в одном предложении.
'''
import docx
import easygui
import os
import sys

def OpenFile():
    # Открывать только docx
    file = easygui.fileopenbox()
    if str(file) != "None":
        finder = file.find("docx")
        if finder == -1:
            file = "None"
    return file

def ReadText():
    openedFile = OpenFile()
    if(str(openedFile) == "None"):
        print("\nERROR: Неверный формат(Нужен только <docx>) / Активация кнопки \"Отмена\"")
        print("Нажмите \"Enter\" для продолжения...")
        input()
        os.system("cls")
        sys.exit()
    doc = docx.Document(openedFile)
    text = ""
    for para in doc.paragraphs:
        text += para.text + " "
    text += "\n"
    return text

def CountWords(text):
    splited = []
    tSplited = text.split()
    i = 0
    while i < len(tSplited):
        check1 = tSplited[i].find('.')
        check2 = tSplited[i].find('!')
        check3 = tSplited[i].find('?')
        if check1 == -1 and check2 == -1 and check3 == -1:
            splited.append(tSplited[i])
        else:
            if (check1 == 0 or check2 == 0 or check3 == 0):
                splited.append(tSplited[i])
            else:
                word = ""
                sumb = ""
                j = 0
                while j < len(tSplited[i])-1:
                    word += tSplited[i][j]
                    j += 1
                sumb = tSplited[i][j]
                splited.append(word)
                splited.append(sumb)
        i += 1
    return splited

def TogetherText(splited):
    k = 0
    complited = ""
    while k < len(splited):
        if (splited[k] == "." or splited[k] == "!" or splited[k] == "?"):
            complited += splited[k]
        else:
            complited += " " + splited[k]
        k += 1
    return complited

def Statistics(splited):
    amountWords = 0
    amountSentences = 0
    i = 0
    while i < len(splited):
        if (splited[i] == "." or splited[i] == "!" or splited[i] == "?"):
            amountSentences += 1
        else:
            amountWords += 1
        i += 1
    
    print("Кол-во всего предложений:", amountSentences)
    print("Кол-во всего слов:", amountWords)

def Main():
    comletedText = "" # Весь текст после обработки и объединения
    splitedText = [] # Разделенный текст на слова и знаки препинания(. ! ?)
    os.system("cls")
    fullText = ReadText()
    splitedText = CountWords(fullText)
    comletedText = TogetherText(splitedText)
    print("Результат:\n")
    print(comletedText, "\n")
    Statistics(splitedText)
    print("\n\nНажмите \"Enter\" для продолжения...")
    input()
    os.system("cls")

if __name__ == '__main__':
    Main()