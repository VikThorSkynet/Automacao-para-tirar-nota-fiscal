import pyautogui
import openpyxl
import time
import webbrowser
import pyperclip

# Carrega o arquivo Excel e seleciona a planilha ativa
# *** IMPORTANTE: Certifique-se que o caminho para o arquivo está correto ***
caminho_arquivo_excel = r"C:\Users\numbe\OneDrive\Área de Trabalho\Importante\NFC Automatico\NF ADULTO.xlsx"
try:
    workbook = openpyxl.load_workbook(caminho_arquivo_excel)
    sheet = workbook.active
except FileNotFoundError:
    print(f"Erro: O arquivo Excel não foi encontrado em '{caminho_arquivo_excel}'. Verifique o caminho.")
    exit() # Encerra o script se o arquivo não for encontrado
except Exception as e:
    print(f"Erro ao carregar o arquivo Excel: {e}")
    exit()

# Define a linha inicial e a quantidade de linhas que deseja percorrer
linha_inicial = 106  # Linha do Excel onde os dados começam (ajuste se necessário)
num_linhas = 1      # Quantidade de linhas a processar a partir da linha_inicial (ajuste conforme necessário)

# URL do formulário
url = "https://www.nfe-cidades.com.br/home/actions/emissaonf2"

print(f"Iniciando processo da linha {linha_inicial} até {linha_inicial + num_linhas - 1}...")

# Loop para percorrer as linhas no Excel
for i in range(linha_inicial, linha_inicial + num_linhas):
    print(f"\nProcessando linha {i}...")

    # --- Leitura dos Dados ---
    # Lê CPF (Coluna C)
    cpf_valor_bruto = sheet[f"C{i}"].value
    # Converte para string e remove espaços extras, se não for None
    cpf = str(cpf_valor_bruto).strip() if cpf_valor_bruto is not None else None

    # Verifica se o campo CPF está vazio ou nulo
    if not cpf:
        print(f"  Linha {i}: CPF (Coluna C) vazio. Pulando para próxima linha.")
        continue

    # Lê Campo B (Coluna D)
    campo_b_valor_bruto = sheet[f"D{i}"].value
    campo_b = str(campo_b_valor_bruto) if campo_b_valor_bruto is not None else "" # Usa string vazia se for None

    # Lê e Formata Campo C (Coluna E) - O VALOR
    valor_bruto_e = sheet[f"E{i}"].value
    campo_c_formatado = None # Inicializa como None

    # Tenta formatar se for um número (int ou float)
    if isinstance(valor_bruto_e, (int, float)):
        try:
            # Formata para 2 casas decimais, usando PONTO como separador decimal temporário
            valor_formatado_temp = f"{valor_bruto_e:.2f}"
            # Substitui o PONTO pela VÍRGULA para o formato brasileiro
            campo_c_formatado = valor_formatado_temp.replace('.', ',')
            print(f"  Linha {i}: Valor (Coluna E) lido como número ({valor_bruto_e}), formatado para '{campo_c_formatado}'.")
        except (TypeError, ValueError):
            # Caso haja algum erro inesperado na formatação, trata como string
            if valor_bruto_e is not None:
                campo_c_formatado = str(valor_bruto_e)
            print(f"  Linha {i}: Valor (Coluna E) é número mas falhou ao formatar. Usando como string: '{campo_c_formatado}'.")

    elif valor_bruto_e is not None:
        # Se não for número (pode já ser texto, como "-"), converte para string
        campo_c_formatado = str(valor_bruto_e)
        print(f"  Linha {i}: Valor (Coluna E) lido como texto: '{campo_c_formatado}'.")
    else:
        print(f"  Linha {i}: Valor (Coluna E) está vazio (None).")
    # Se valor_bruto_e for None, campo_c_formatado permanecerá None

    # --- Verificação de Pular Linha ---
    # Verifica a cor da célula E
    cor_celula_fill = sheet[f"E{i}"].fill
    cor_celula_rgb = None
    is_yellow = False
    if cor_celula_fill and cor_celula_fill.start_color and cor_celula_fill.start_color.rgb:
         cor_celula_rgb = cor_celula_fill.start_color.rgb
         # Verifica os 6 últimos caracteres (RGB) para Amarelo (FFFF00), ignorando Alpha (FF inicial opcional)
         if cor_celula_rgb[-6:] == "FFFF00":
             is_yellow = True
             print(f"  Linha {i}: Célula E está amarela.")

    # Verifica se o valor formatado está vazio, contém "-" ou a célula é amarela
    if not campo_c_formatado or campo_c_formatado == "-" or is_yellow:
        print(f"  Linha {i}: Valor inválido ('{campo_c_formatado}') ou célula amarela. Pulando para próxima linha.")
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

    # Rola a página para baixo para visualizar os campos
    pyautogui.scroll(-800)
    print("  Página rolada para baixo.")
    time.sleep(2)

    # Foco e preenchimento do campo CPF
    try:
        pyautogui.click(x=499, y=215) # Ajuste coordenadas se necessário
        time.sleep(1)
        pyautogui.write(cpf, interval=0.1) # Usa a variável cpf já tratada
        print(f"  CPF '{cpf}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com campo CPF: {e}")
        pyautogui.hotkey('ctrl', 'w') # Tenta fechar a aba em caso de erro
        continue # Pula para próxima linha

    time.sleep(2)
    pyautogui.scroll(-500) # Rola mais um pouco se necessário
    time.sleep(1)

    # Foco e preenchimento do segundo campo (Descrição?)
    try:
        pyautogui.click(x=529, y=323) # Ajuste coordenadas se necessário
        time.sleep(1)
        pyautogui.write(campo_b) # Usa a variável campo_b já tratada
        print(f"  Campo B '{campo_b}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com segundo campo: {e}")
        pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(2)

    # Foco e preenchimento do terceiro campo (Valor)
    try:
        pyautogui.click(x=1281, y=300) # Ajuste coordenadas se necessário
        time.sleep(1)
        # Usa a string formatada campo_c_formatado
        pyautogui.write(campo_c_formatado, interval=0.1)
        print(f"  Campo C (Valor) '{campo_c_formatado}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com terceiro campo (Valor): {e}")
        pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(2)
    pyautogui.scroll(-600) # Rola para visualizar o botão emitir
    time.sleep(1)

    # Clique no botão Emitir
    try:
        pyautogui.click(x=1137, y=495) # Ajuste coordenadas se necessário
        print("  Botão 'Emitir' clicado.")
    except Exception as e:
        print(f"  Erro ao clicar no botão 'Emitir': {e}")
        pyautogui.hotkey('ctrl', 'w')
        continue

    time.sleep(2) # Espera alguma resposta da página

    # Seleção e cópia de texto (AJUSTE AS COORDENADAS E O MOVIMENTO!)
    print("  Tentando selecionar e copiar texto de resposta...")
    try:
        pyautogui.moveTo(x=512, y=512) # Posição inicial da seleção
        time.sleep(1) # Reduzi a pausa
        pyautogui.mouseDown()
        time.sleep(0.5) # Pausa curta com botão pressionado
        # Ajuste o valor 45 se a palavra for maior ou menor
        pyautogui.move(45, 0) # Move mais suavemente
        time.sleep(0.5) # Pausa curta antes de soltar
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


    # Salva o valor na coluna F do Excel
    try:
        coluna_destino = f"F{i}"
        sheet[coluna_destino] = texto_copiado
        workbook.save(caminho_arquivo_excel)
        print(f"  Texto copiado salvo na célula {coluna_destino}.")
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