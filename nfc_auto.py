import pyautogui
import openpyxl
import time
import webbrowser
import win32clipboard

# Carrega o arquivo Excel e seleciona a planilha ativa
workbook = openpyxl.load_workbook("C:\\Users\\Number One\\Desktop\\Importante\\NFC Automatico\\teste.xlsx")
sheet = workbook.active

# Define a quantidade de linhas que deseja percorrer
num_linhas = 24  # Exemplo, ajuste para o número de linhas necessárias

# URL do formulário
url = "https://www.nfe-cidades.com.br/home/actions/emissaonf2"

# Loop para percorrer as linhas no Excel (ajuste o intervalo conforme necessário)
for i in range(1, num_linhas + 1):
    # Lê os valores das células A, B, e C da linha atual
    cpf = sheet[f"A{i}"].value
    campo_b = sheet[f"B{i}"].value
    campo_c = sheet[f"C{i}"].value

    # Pausa para ajuste (garante tempo para a página carregar ou o usuário posicionar o cursor)
    time.sleep(3)

    webbrowser.open(url)

    time.sleep(10)

    # Rola a página para baixo
    pyautogui.scroll(-800)  # Valor negativo para rolar para baixo; positivo para cima
    
    time.sleep(2)

    # Foco no campo CPF na página da web e insere o valor de 'cpf'
    pyautogui.click(x=499, y=215)  # Coordenada do campo CPF
    time.sleep(1)
    pyautogui.write(str(cpf), interval=0.1)

    # Pausa para mudar para o próximo campo
    time.sleep(2)

    pyautogui.scroll(-500)

    time.sleep(1)

    # Foco no segundo campo e insere o valor de 'campo_b'
    pyautogui.click(x=529, y=323)  # Coordenada do segundo campo
    time.sleep(1)
    pyautogui.write(str(campo_b))

    # Pausa para mudar para o terceiro campo
    time.sleep(2)

    # Foco no terceiro campo e insere o valor de 'campo_c'
    pyautogui.click(x=1281, y=300)  # Coordenada do terceiro campo
    time.sleep(1)
    pyautogui.write(str(campo_c), interval=0.1)

    # Pausa entre linhas para evitar sobrecarga e permitir revisão
    time.sleep(2)

    pyautogui.scroll(-600)

    time.sleep(1)

    pyautogui.click(x=1137, y=495) #Click emitir

    pyautogui.moveTo(x=513, y=429)

    time.sleep(5)

    # Seleciona a palavra com clique e arraste
    #pyautogui.moveTo(x=513, y=429)
    #time.sleep(4)
    #pyautogui.mouseDown()      # Pressiona o botão do mouse
    #time.sleep(3)            # Pequena pausa
    #pyautogui.move(40, 0)    # Move o mouse x pixels para a direita para selecionar a palavra
    #time.sleep(3)
    #pyautogui.mouseUp()        # Solta o botão do mouse 

    #time.sleep(3)

    # Copia o texto selecionado usando Ctrl + C
    #pyautogui.hotkey("ctrl", "c")

    #time.sleep(3)

    pyautogui.hotkey('ctrl', 'w')

    #time.sleep(3)

    #win32clipboard.OpenClipboard()
    #texto_copiado = win32clipboard.GetClipboardData()
    #win32clipboard.CloseClipboard()
    
    # Salva o valor na coluna D do Excel
    #sheet[f"D{i}"] = texto_copiado
    #workbook.save("C:\\Users\\Number One\\Desktop\\Importante\\NFC Automatico\\teste.xlsx")



print("Processo concluído!")


