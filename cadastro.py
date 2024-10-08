import pandas as pd
import time
import pyperclip
import pyautogui
import pygetwindow as gw

# Verificar se o Total On está em foco
def garantir_foco_tota_on():
    try:
        window = gw.getWindowsWithTitle('Total On')[0]
        if window.isMinimized:
            window.restore()
        window.activate()
        time.sleep(1)  # Aguarde para garantir que a janela está ativa
    except IndexError:
        print("A janela do Total On não foi encontrada.")

def selecionar_texto(x_inicio, y_inicio, x_fim, y_fim):
    # Move o mouse para a posição inicial e pressiona o botão esquerdo
    pyautogui.moveTo(x_inicio, y_inicio)
    pyautogui.mouseDown(button='left')
    
    # Aguarda um pouco para garantir que o botão esteja pressionado
    time.sleep(0.1)
    
    # Move o mouse até a posição final
    pyautogui.moveTo(x_fim, y_fim)
    
    # Solta o botão do mouse
    pyautogui.mouseUp()

# Lendo um arquivo .xlsm
df = pd.read_excel('vambora.xlsx', sheet_name='Planilha1')
codbarras = df['CÓD BARRAS'].astype(str).str.replace('.0', '', regex=False).tolist()
descricao = df['DESCRIÇÃO'].tolist()
st = df['ST'].tolist()
sku = df['REF'].tolist()
custoNormal = df['CUSTO'].tolist()
custo = [str(round(valor, 5)).replace('.', ',') for valor in custoNormal]
vendaNormal = df['VENDA'].tolist()
venda = [str(round(valor, 5)).replace('.', ',') for valor in vendaNormal]
atacadoNormal = df['ATACADO'].tolist()
atacado = [str(round(valor, 5)).replace('.', ',') for valor in atacadoNormal]

i = 0
u = 0
d = 0
cst = 0
om = 0
#pyautogui.hotkey('alt', 'tab') # abrir o total On
garantir_foco_tota_on()
while i < 29:                # ALTERE PARA QUANTIDADE DE PRODUTOS NA NOTA
    pyautogui.click(x=600 ,y=388) # abrir dados cadastrais
    pyautogui.doubleClick(x=523 ,y=354) # selecionar todo o campo de codigo barras
    pyautogui.press("backspace") # apagar se tiver alguma coisa
    time.sleep(2)
    #pyautogui.click(x=523 , y=354) #clicar no parametro de codigo de barras
    pyperclip.copy(codbarras[i]) # copia o codigo de barras na posição i para input do teclado
    pyautogui.hotkey("ctrl", "v") # cola
    pyautogui.press("enter") # pressiona enter para pesquisar produto
    time.sleep(15)

    # parte de dados cadastrais do produto
    selecionar_texto(966, 354, 629, 354) # seleciona descrição do produto
    pyperclip.copy(descricao[i])
    pyautogui.hotkey("ctrl", "v")

    selecionar_texto(657, 444, 573, 444) # seleciona conteudo do SKU
    time.sleep(1)
    pyperclip.copy(sku[i]) # Copia valor na posição i do array
    pyautogui.hotkey("ctrl", "v") # cola
    pyautogui.press("enter")

    pyautogui.doubleClick(x=614 ,y=685) # seleciona conteudo do preço de custo
    pyautogui.press("backspace") # apaga conteudo
    time.sleep(1)
    pyautogui.click(x=614 ,y=685) # clica no preço de custo
    pyperclip.copy(custo[i]) # Copia valor na posição i do array
    pyautogui.hotkey("ctrl", "v") # cola
    pyautogui.press("enter")

    pyautogui.doubleClick(x=759 ,y=685) # seleciona conteudo do preço de custo
    pyautogui.press("backspace") # apaga conteudo
    time.sleep(1)
    pyautogui.click(x=759 ,y=685) # clica no preço de venda
    pyperclip.copy(venda[i]) # Copia valor na posição i do array
    pyautogui.hotkey("ctrl", "v") # cola
    pyautogui.press("enter")

    pyautogui.doubleClick(x=541 ,y=755) # seleciona conteudo do preço de custo
    pyautogui.press("backspace") # apaga conteudo
    time.sleep(1)
    pyautogui.click(x=541 ,y=755) # clica no atacado
    pyautogui.rightClick(x=541 ,y=755)
    pyperclip.copy(atacado[i]) # Copia valor na posição i do array
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press("enter")

    pyautogui.click(x=1011 ,y=502)
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.click(x=1011 ,y=502)

    pyautogui.click(x=1097 ,y=500)
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.click(x=1097 ,y=500)

    # parte de tributações
    pyautogui.click(x=829 ,y=391) # abre a opção de tributações
    pyautogui.click(x=873 ,y=356) # clica no nome do produto e da enter
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.doubleClick(x=976 ,y=430) # clica no CFOP para verificar se tem ou não tem ST
    pyautogui.rightClick(x=976 ,y=430) # botão direto
    if (st[i + 1] == 0.0): 
        res = pyperclip.copy("5102")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.click(x=1310 ,y=571)
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.click(x=1310, y=541) 
        pyautogui.click(x=674 ,y=502) # clica em situação tributaria
        while (cst <= 14):
            pyautogui.press("up")
            cst += 1
        pyautogui.click(x=916 ,y=778) # fecha situação tributaria
        cst = 0
        time.sleep(1)
    elif (st[i] != 0.0):
        result = pyperclip.copy("5405")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.click(x=674 ,y=502) # clica em situação tributaria
        while (cst <= 14):
            pyautogui.press("down")
            cst += 1
        cst = 0
        while (cst < 3):
            pyautogui.press("up")
            cst += 1
        pyautogui.press('enter')
        pyautogui.click(x=1304, y=587)
        cst = 0
    else:
        print("ERROR!")
        exit()
    pyautogui.click(x=1276 ,y=596) # clica em aliquota do ICMS
    while (u <= 49):
        pyautogui.press("up") # sobe 50 vezes
        u += 1
    while (d < 19):
        pyautogui.press("down") # desce 20 vezes
        d += 1
    u = 0
    d = 0
    pyautogui.click(x=916 ,y=778)
    pyautogui.click(x=717 ,y=426) # seleciona origem de mercado
    while (om < 8):
        pyautogui.press("up")
        om += 1
    om = 0
    time.sleep(1)
    pyautogui.click(x=600 ,y=388) # volta para os dados cadastrais do produto
    pyautogui.click(x=612 ,y=221) # salvar registro
    pyautogui.press('enter')
    pyautogui.press('enter')

    i += 1
