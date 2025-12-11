import pyautogui
import openpyxl
import time
import webbrowser
import pyperclip

# *** IMPORTANTE: Certifique-se que o caminho para o arquivo está correto ***
# Definição dos 2 caminhos de arquivos Excel disponíveis
caminho_excel_1 = r"C:\\Users\\numbe\\OneDrive\\Área de Trabalho\\Importante\\NFC Automatico\\NF ORIGEM.xlsx"
caminho_excel_2 = r"C:\\Users\\numbe\\OneDrive\\Área de Trabalho\\Importante\\NFC Automatico\\NF ADULTO.xlsx"

# Pergunta ao usuário qual arquivo deseja usar
print("Escolha qual arquivo Excel deseja usar:")
print("1 - ORIGEM")
print("2 - ADULTO")
escolha = input("Digite 1 ou 2: ").strip()

# Seleciona o caminho baseado na escolha do usuário
if escolha == "1":
    caminho_arquivo_excel = caminho_excel_1
    print(f"Arquivo selecionado: ORIGEM")
elif escolha == "2":
    caminho_arquivo_excel = caminho_excel_2
    print(f"Arquivo selecionado: ADULTO")
else:
    print("Escolha inválida!")
    exit()

try:
    workbook = openpyxl.load_workbook(caminho_arquivo_excel)
    sheet = workbook.active
    print(f"Arquivo Excel carregado com sucesso: {caminho_arquivo_excel}")
except FileNotFoundError:
    print(f"Erro: O arquivo Excel não foi encontrado em '{caminho_arquivo_excel}'. Verifique o caminho.")
    print("Verifique se o caminho está correto e se o arquivo existe.")
    exit() # Encerra o script se o arquivo não for encontrado
except Exception as e:
    print(f"Erro ao carregar o arquivo Excel: {e}")
    exit()

# Lê Campo G10 (Coluna G10) 
# Define a linha inicial. Linha do Excel onde os dados começam
campo_G10 = sheet[f"G10"].value
print(f"Linha do Excel onde os dados começam {campo_G10}")   

# Lê Campo G13 (Coluna G13) 
# Quantidade de linhas a processar a partir da linha_inicial
campo_G13 = sheet[f"G13"].value
print(f"Linha do Excel onde os dados começam {campo_G13}")  

# Define a linha inicial e a quantidade de linhas que deseja percorrer
linha_inicial = campo_G10  # Linha do Excel onde os dados começam (ajuste se necessário)
num_linhas = campo_G13    # Quantidade de linhas a processar a partir da linha_inicial (ajuste conforme necessário)

# URL do formulário
url = "https://contribuinte.nota-eletronica.betha.cloud/#/ZGF0YWJhc2U6NDY2LGVudGl0eTo4MzQsbm90YV9lbGV0cm9uaWNhX2NvbnRyaWJ1aW50ZTo1MDE2MzY0/notas-fiscais/nota-emissao"

print(f"Iniciando processo da linha {linha_inicial} até {linha_inicial + num_linhas - 1}...")

# Loop para percorrer as linhas no Excel
for i in range(linha_inicial, linha_inicial + num_linhas):
    print(f"\nProcessando linha {i}...")

    # Verifica se a coluna E já está preenchida (pula se houver algo)
    valor_coluna_e = sheet[f"E{i}"].value
    if valor_coluna_e is not None and str(valor_coluna_e).strip() != "":
        print(f"  Linha {i}: Coluna E já preenchida. Pulando para próxima linha.")
        continue

    # --- Leitura dos Dados ---
    # Lê CPF (Coluna C)
    cpf_valor_bruto = sheet[f"C{i}"].value
    # Converte para string, remove formatação e deixa apenas números
    if cpf_valor_bruto is not None:
        cpf = ''.join(filter(str.isdigit, str(cpf_valor_bruto)))
    else:
        cpf = None

    # Verifica se o campo CPF está vazio ou nulo
    if not cpf:
        print(f"  Linha {i}: CPF (Coluna C) vazio. Pulando para próxima linha.")
        continue

    # Lê Campo G2 (Coluna G2) "Total de tributos"
    campo_G2_valor_bruto = sheet[f"G2"].value
    campo_G2 = str(campo_G2_valor_bruto) if campo_G2_valor_bruto is not None else "" # Usa string vazia se for None
    print(f"Total de tributos: {campo_G2}")
    
    # Lê Campo G4 (Coluna G4) "Alíguota"
    campo_G4_valor_bruto = sheet[f"G4"].value
    if campo_G4_valor_bruto is not None:
        # Converte para string e substitui ponto por vírgula para manter formato brasileiro
        campo_G4 = str(campo_G4_valor_bruto).replace('.', ',')
    else:
        campo_G4 = ""  # Usa string vazia se for None
    print(f"Alíguota: {campo_G4}")
    
    # Lê Campo G7 (Coluna G7) Data para tirar as NF
    campo_G7_valor_bruto = sheet[f"G7"].value
    if campo_G7_valor_bruto is not None:
        # Verifica se é um objeto datetime
        if hasattr(campo_G7_valor_bruto, 'strftime'):
            # Formata para DD/MM/YYYY
            campo_G7 = campo_G7_valor_bruto.strftime('%d/%m/%Y')
        else:
            # Se não for datetime, converte para string
            campo_G7 = str(campo_G7_valor_bruto)
    else:
        campo_G7 = "" # Usa string vazia se for None
    print(f"Data para tirar Nf {campo_G7}")

    # Lê e Formata Campo VALOR (Coluna D) - O VALOR
    valor_bruto_D = sheet[f"D{i}"].value
    campo_D_formatado = None # Inicializa como None

    # Tenta formatar se for um número (int ou float)
    if isinstance(valor_bruto_D, (int, float)):
        try:
            # Formata para 2 casas decimais, usando PONTO como separador decimal temporário
            valor_formatado_temp = f"{valor_bruto_D:.2f}"
            # Substitui o PONTO pela VÍRGULA para o formato brasileiro
            campo_D_formatado = valor_formatado_temp.replace('.', ',')
            print(f"  Linha {i}: Valor (Coluna D) lido como número ({valor_bruto_D}), formatado para '{campo_D_formatado}'.")
        except (TypeError, ValueError):
            # Caso haja algum erro inesperado na formatação, trata como string
            if valor_bruto_D is not None:
                campo_D_formatado = str(valor_bruto_D)
            print(f"  Linha {i}: Valor (Coluna D) é número mas falhou ao formatar. Usando como string: '{campo_D_formatado}'.")

    elif valor_bruto_D is not None:
        # Se não for número (pode já ser texto, como "-"), converte para string
        campo_D_formatado = str(valor_bruto_D)
        print(f"  Linha {i}: Valor (Coluna D) lido como texto: '{campo_D_formatado}'.")
    else:
        print(f"  Linha {i}: Valor (Coluna D) está vazio (None).")
    # Se valor_bruto_D for None, campo_D_formatado permanecerá None

    # --- Verificação de Pular Linha ---
    # Verifica a cor da célula D
    cor_celula_fill = sheet[f"D{i}"].fill
    cor_celula_rgb = None
    is_yellow = False
    if cor_celula_fill and cor_celula_fill.start_color and cor_celula_fill.start_color.rgb:
         cor_celula_rgb = cor_celula_fill.start_color.rgb
         # Converte o objeto RGB para string para verificar se é amarelo
         cor_rgb_str = str(cor_celula_rgb)
         # Verifica se contém FFFF00 (amarelo) na string RGB
         if "FFFF00" in cor_rgb_str:
             is_yellow = True
             print(f"  Linha {i}: Célula D está amarela.")

    # Verifica se o valor formatado está vazio, contém "-" ou a célula é amarela
    if not campo_D_formatado or campo_D_formatado == "-" or is_yellow:
        print(f"  Linha {i}: Valor inválido ('{campo_D_formatado}') ou célula amarela. Pulando para próxima linha.")
        continue

    # --- Automação Web ---
    print(f"  Linha {i}: Dados válidos. Iniciando automação web...")
    # Pausa antes de abrir o navegador
    time.sleep(3)

    try:
        webbrowser.open(url)
        print(f"  Linha {i}: Navegador aberto com URL: {url}")
    except Exception as e:
        print(f"  Erro ao abrir o navegador ou URL: {e}")
        continue # Pula para a próxima linha se não conseguir abrir

    # Aumenta a espera para garantir o carregamento completo da página
    time.sleep(10) # Espera 10 segundos

    #Se for retroativo use o código abaixo<<<<<<<<<<< SE NÃO FOR USAR FAVOR COMENTAR ESSE TRECHO DE CODIGO

    pyautogui.click(x=112, y=383) # Posição inicial da seleção
    time.sleep(1.5) # Reduzi a pausa

    #coloque a data desejada
    pyautogui.write(campo_G7) # DATA PARA TIRAR NF
    time.sleep(1)
    print("Data colocada")

    # Foco e preenchimento do campo CPF
    try:
        pyautogui.doubleClick(x=305, y=538) # Ajuste coordenadas se necessário
        time.sleep(1)
        pyautogui.write(cpf, interval=0.1) # Usa a variável cpf já tratada
        print(f"  CPF '{cpf}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com campo CPF: {e}")
        # pyautogui.hotkey('ctrl', 'w') # Tenta fechar a aba em caso de erro
        continue # Pula para próxima linha

    time.sleep(1.2)
    pyautogui.doubleClick(x=416, y=585) # Click no nome que aparece embaixo do cpf
    time.sleep(1)
    pyautogui.scroll(-1400) # Rola para visualizar o botão emitir
    time.sleep(1)

    #Preenchendo ALÍGUOTA
    pyautogui.doubleClick(x=970, y=520)
    time.sleep(1)
    pyautogui.doubleClick(x=1010, y=600)
    time.sleep(1)
    pyautogui.doubleClick(x=1175, y=525)
    time.sleep(1)
    pyautogui.write(campo_G4)
    time.sleep(1)

    # Foco e preenchimento do segundo campo (Total de tributos a serem pagos)
    try:
        pyautogui.doubleClick(x=115, y=595) # Ajuste coordenadas se necessário
        time.sleep(1)
        pyautogui.write(campo_G2) # Usa a variável campo_G2 já tratada
        print(f"  Campo G2 '{campo_G2}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com segundo campo: {e}")
        # pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(2)

    # Foco e preenchimento do terceiro campo (Valor)
    try:
        pyautogui.moveTo(x=505, y=455)
        time.sleep(1)
        pyautogui.scroll(-250)
        time.sleep(1)
        pyautogui.doubleClick(x=153, y=475) # Ajuste coordenadas se necessário
        time.sleep(1)
        # Usa a string formatada campo_D_formatado
        pyautogui.write(campo_D_formatado, interval=0.1)
        print(f"  Campo D (Valor) '{campo_D_formatado}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com terceiro campo (Valor): {e}")
        # pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(1.2)

    # Clique no botão Emitir
    try:
        pyautogui.doubleClick(x=903, y=688) # Ajuste coordenadas se necessário
        print("  Botão 'Emitir' clicado.")
    except Exception as e:
        print(f"  Erro ao clicar no botão 'Emitir': {e}")
        pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(15) # Espera alguma resposta da página

    # Seleção e cópia de texto (AJUSTE AS COORDENADAS E O MOVIMENTO!)
    print("  Tentando selecionar e copiar texto de resposta...")
    try:
        pyautogui.moveTo(x=109, y=315) # Posição inicial da seleção
        time.sleep(1) # Reduzi a pausa
        pyautogui.mouseDown()
        time.sleep(1) # Pausa curta com botão pressionado
        # Ajuste o valor 45 se a palavra for maior ou menor
        pyautogui.move(115, 0) # Move mais suavemente
        time.sleep(1) # Pausa curta antes de soltar
        pyautogui.mouseUp()
        print("  Texto selecionado (visualmente).")
        time.sleep(1)

        # Copia o texto selecionado
        pyautogui.hotkey("ctrl", "c")
        print("  Comando Ctrl+C enviado.")
        time.sleep(1) # Espera a cópia acontecer

    except Exception as e:
        print(f"  Erro durante seleção/cópia do texto: {e}")
        # Mesmo com erro aqui, tenta ler o clipboard e fechar a aba

    # Fecha a aba/janela do navegador
    pyautogui.hotkey('ctrl', 'w')
    print("  Comando Ctrl+W enviado para fechar aba/janela.")
    time.sleep(2) # Espera fechar

    # --- Leitura da Área de Transferência e Salvamento ---
    texto_copiado = "Erro Leitura Clipboard" # Valor padrão em caso de falha
    try:
        texto_copiado = pyperclip.paste()
        # Limpa espaços extras que podem vir na cópia
        texto_copiado = texto_copiado.strip() if texto_copiado else "Clipboard Vazio"
        print(f"  Texto copiado da área de transferência: '{texto_copiado}'")
    except pyperclip.PyperclipException as e: # Captura exceção específica do pyperclip
        print(f"  Erro ao ler da área de transferência (pyperclip): {e}")
    except Exception as e: # Captura outras exceções genéricas
         print(f"  Erro inesperado ao ler da área de transferência: {e}")


    # Salva o valor na coluna E do Excel
    try:
        coluna_destino = f"E{i}"
        # Força o Excel a tratar como texto adicionando aspas simples no início
        sheet[coluna_destino] = f"'{texto_copiado}"
        workbook.save(caminho_arquivo_excel)
        print(f"  Texto copiado salvo na célula {coluna_destino} como texto: '{texto_copiado}'")
    except PermissionError:
        print(f"ERRO CRÍTICO: Não foi possível salvar o arquivo '{caminho_arquivo_excel}'.")
        print("Verifique se o arquivo está aberto em outro programa ou se você tem permissão de escrita.")
        print("O script será interrompido para evitar perda de dados.")
        break # Interrompe o loop principal se não conseguir salvar
    except Exception as e:
        print(f"  Erro inesperado ao salvar o arquivo Excel: {e}")
        print("  Tentando continuar para a próxima linha, mas o salvamento falhou.")
        # Decide se quer parar ou continuar. Parar é mais seguro.
        break


# Mensagem final fora do loop
print("\nProcesso concluído!")

