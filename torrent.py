# -*- coding: utf-8 -*-
import xlrd
import time
import re
import pyautogui

book = xlrd.open_workbook("torrent.xlsx")
sh = book.sheet_by_index(0)
x = (sh.nrows)
print("Quantidade de animes: " + str(x))
print("Iniciando em tres segundos")
time.sleep(3)
pyautogui.click(171, 751)

for rx in range(0, x):
    link = sh.cell_value(rowx=rx, colx=3)
    nome = sh.cell_value(rowx=rx, colx=0)
    nome = nome.replace("[C]","")
    nome = nome.replace("[O]","")
    nome = nome.replace("[E]","")
    nome = nome.replace("[M]","")
    nome = nome.replace(" + ","")
    nome = nome.replace("+","")
    nome = nome.replace(" ","+")
    nome = nome.replace("++","+")

    tamanho = sh.cell_value(rowx=rx, colx=1)
    fanssuber = sh.cell_value(rowx=rx, colx=2)
    id = re.search(r'(?<=id=)\w+', link)
    id = id.group(0)
    DownloadPage = "http://fansubber.net/tracker/index.php?page=downloadcheck&id=" + id
    #http://fansubber.net/tracker/download.php?id=5d44447bf9c4ca27be7de93b9729a6d2af475650&f=%5BC%5D+%5BO%5D+Rurouni+Kenshin+Shin+Kyoto-Hen+%5BBD%7ERip%2C+1080p%2C+10bits%5D.torrent&key=fee4bc17
    DownloadArq = "fansubber.net/tracker/download.php?id=" + id + "&f=" + nome + ".torrent&key=fee4bc17"
    print("link  Fansubber:" + link )
    print("Pagina de download:" + DownloadPage)
    print("Link de Download : " + DownloadArq)
    print("id:" + id)
    print("Tamanho:"+ tamanho)
    print("nome:" + nome )

    print("============================================================================================")
    print("")
    time.sleep(1)
    pyautogui.click(573, 106)
    pyautogui.click(573, 106)
    time.sleep(1)
    pyautogui.typewrite(str(DownloadArq))
    time.sleep(1)
    pyautogui.press("return")
    time.sleep(4)



    # tudo print(sh.row(rx))
