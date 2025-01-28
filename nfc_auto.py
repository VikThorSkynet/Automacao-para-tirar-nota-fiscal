import pyautogui
import openpyxl
import time
import webbrowser
import win32clipboard
import pyperclip

# Carrega o arquivo Excel e seleciona a planilha ativa
workbook = openpyxl.load_workbook(r"C:\Users\numbe\OneDrive\Área de Trabalho\Importante\NFC Automatico\teste.xlsx")
sheet = workbook.active

# Define a quantidade de linhas que deseja percorrer
num_linhas = 17  # Exemplo, ajuste para o número de linhas necessárias

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

    time.sleep(4)

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
    # Garante que o valor seja tratado como string e mantenha os zeros
    campo_c_formatado = str(campo_c).zfill(len(str(campo_c)))  # Mantém os zeros à direita
    pyautogui.write(campo_c_formatado, interval=0.1)

    # Pausa entre linhas para evitar sobrecarga e permitir revisão
    time.sleep(2)

    pyautogui.scroll(-600)

    time.sleep(1)

    pyautogui.click(x=1137, y=495) #Click emitir

    #pyautogui.moveTo(x=512, y=512)

    time.sleep(2)

    # Seleciona a palavra com clique e arraste
    pyautogui.moveTo(x=512, y=512)
    time.sleep(2)
    pyautogui.mouseDown()      # Pressiona o botão do mouse
    time.sleep(2)            # Pequena pausa
    pyautogui.move(34, 0)    # Move o mouse x pixels para a direita para selecionar a palavra
    time.sleep(2)
    pyautogui.mouseUp()        # Solta o botão do mouse 

    time.sleep(2)

    # Copia o texto selecionado usando Ctrl + C
    pyautogui.hotkey("ctrl", "c")

    time.sleep(2)

    pyautogui.hotkey('ctrl', 'w')

    time.sleep(2)

    # Lê o conteúdo da área de transferência
    try:
        texto_copiado = pyperclip.paste()
    except:
        texto_copiado = "Erro na leitura"
        print("Erro ao ler da área de transferência")
    
    # Salva o valor na coluna D do Excel
    sheet[f"D{i}"] = texto_copiado
    try:
        workbook.save(r"C:\\Users\\numbe\\OneDrive\\Área de Trabalho\\Importante\\NFC Automatico\\teste.xlsx")
    except PermissionError:
        print("Erro ao salvar: Feche o arquivo Excel e tente novamente.")
        break

print("Processo concluído!")


