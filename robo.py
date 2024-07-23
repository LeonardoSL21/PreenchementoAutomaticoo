import pyautogui
import time
import pandas as pd


# Entar no excel
pyautogui.press('win')
pyautogui.write('microsoft edge')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(309, 44)
pyautogui.write('https://www.microsoft365.com/launch/excel?auth=1')
pyautogui.press('enter')
time.sleep(4)
pyautogui.write('leocg20sl@gmail.com')
pyautogui.press('enter')
time.sleep(2)
pyautogui.write('le290498')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(39, 422)
time.sleep(2)
pyautogui.click(268, 257)
time.sleep(5)

# importar base de dados
produtos = pd.read_csv("produtos.csv")
print(produtos)
# cadastrar o produto
#codigo
pyautogui.click(80, 225)
pyautogui.write('codigo')

# repertir o processo até finalizar os produtos da lista
for item in produtos.index:
    # codigo
    codigo = str(produtos.loc[item, 'codigo'])
    pyautogui.write(codigo)

    #marca
    pyautogui.press('right')
    marcar = str(produtos.loc[item, 'marca'])
    pyautogui.write(marcar)

    #tipo
    pyautogui.press('right')
    tipo = str(produtos.loc[item, 'tipo'])
    pyautogui.write(tipo)

    #categoria
    pyautogui.press('right')
    categoria =str(produtos.loc[item, 'categoria'])
    pyautogui.write(categoria)

    #preco_unitario
    pyautogui.press('right')
    preco_unitario = str(produtos.loc[item, 'preco_unitario'])
    pyautogui.write(preco_unitario)

    #cust
    pyautogui.press('right')
    custo = str(produtos.loc[item, 'custo'])
    pyautogui.write(custo)

    #obs
    pyautogui.press('right')
    obs = str(produtos.loc[item,'obs'])
    pyautogui.write(obs)

    i = 0
    while i < 6:
        pyautogui.press('left')
        i +=1
    pyautogui.press('down')


