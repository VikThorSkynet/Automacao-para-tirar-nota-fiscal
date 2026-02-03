# -*- coding: utf-8 -*-
import pyautogui
import openpyxl
import time
import webbrowser
import pyperclip
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

 # Solicita os caminhos dos arquivos
messagebox.showinfo("Informação", "Escolha o arquivo Excel.")
caminho_arquivo_excel = filedialog.askopenfilename(
title="Selecione um Arquivo",
filetypes=(("Planilhas Excel", "*.xlsx;*.xls"), ("Todos os Arquivos", "*.*"))
)

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
url = "https://contribuinte.nota-eletronica.betha.cloud/#/ZGF0YWJhc2U6NDY2LGVudGl0eTo4MzQsbm90YV9lbGV0cm9uaWNhX2NvbnRyaWJ1aW50ZTo1MDE2MzY0/notas-fiscais/dps-emissao"

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

    #Posição data de emissão
    pyautogui.click(x=196, y=466) # Posição inicial da seleção
    time.sleep(1.5) # Reduzi a pausa

    #coloque a data desejada
    pyautogui.write(campo_G7) # DATA PARA TIRAR NF
    time.sleep(1)
    print("Data colocada")
    time.sleep(1)
    pyautogui.click(x=196, y=600)
    time.sleep(1)

    #Preenchimento d0o campo CPF
    try:
        pyautogui.doubleClick(x=165, y=645) # Ajuste coordenadas se necessário
        time.sleep(1.5)
        pyautogui.scroll(-100)
        time.sleep(1.5)
        pyautogui.doubleClick(x=171, y=640)
        time.sleep(1)
        pyautogui.write(cpf, interval=0.1) # Usa a variável cpf já tratada
        print(f"  CPF '{cpf}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com campo CPF: {e}")
        # pyautogui.hotkey('ctrl', 'w') # Tenta fechar a aba em caso de erro
        continue # Pula para próxima linha

    time.sleep(1.2)
    pyautogui.doubleClick(x=219, y=687) # Click no nome que aparece embaixo do cpf
    time.sleep(1)

    #RESOLVER PROBLEMA ENDEREÇO
    pyautogui.scroll(-300)
    time.sleep(1)
    pyautogui.click(x=163, y=493)
    time.sleep(1)
    pyautogui.click(x=226, y=551)
    time.sleep(1)
    pyautogui.write("0")
    time.sleep(1)
    
    #Posição do botão avançar
    pyautogui.scroll(-500)
    time.sleep(1)
    pyautogui.click(x=1210, y=642)
    time.sleep(2)
    print("clicou no botão Avançar")

    #PAGINA 2 SERVIÇO

    #Pais de prestação
    pyautogui.scroll(1000)
    time.sleep(1)
    pyautogui.click(x=189, y=501)
    time.sleep(1)
    pyautogui.write("Brasil")
    time.sleep(1)
    pyautogui.click(x=219, y=539)
    print("Digitou Brasil no campo: Pais de prestação")

    #Municipio de prestação
    time.sleep(1)
    pyautogui.click(x=572, y=500)
    time.sleep(1)
    pyautogui.write("Lagoa Santa")
    time.sleep(1)
    pyautogui.click(x=591, y=599)
    time.sleep(1)
    print("Digitou Lagoa Santa no campo: Municipio de prestação")

    #Natureza da operação
    pyautogui.click(x=943, y=504)
    time.sleep(1)
    pyautogui.write("Op")
    time.sleep(1)
    pyautogui.click(x=960, y=468)
    time.sleep(1)
    print("Digitou Op no campo: Natureza da operação")

    #Tipo de retenção ISSQN
    pyautogui.click(x=225, y=570)
    time.sleep(1)
    pyautogui.click(x=240, y=604)
    time.sleep(1)
    print("Tipo de retenção ISSQN - Clickou em Não retido")

    #Lista de Serviço
    pyautogui.click(x=184, y=632)
    time.sleep(1.2)
    pyautogui.write("08.02.01")
    time.sleep(1.2)
    pyautogui.click(x=289, y=586)
    time.sleep(1)
    pyautogui.click(x=610, y=586)
    time.sleep(1.3)
    print("Selecionou - 08.02.01 - Instrução, treinamento, orientação pedagógica e educacional, avaliação de conhecimentos de qualquer natureza.")

    #NBS
    pyautogui.scroll(-300)
    time.sleep(1)
    pyautogui.click(x=247, y=450)
    time.sleep(1.2)
    pyautogui.write("122051300 - Serviços de educação em línguas estrangeiras e de sinais")
    time.sleep(1.2)
    pyautogui.click(x=329, y=524)
    time.sleep(1)
    print("122051300 - Serviços de educação em línguas estrangeiras e de sinais")

    #Descrição detalhada do serviço
    pyautogui.scroll(-300)
    time.sleep(1)
    pyautogui.click(x=177, y=340)
    time.sleep(1)

     # Copia o campo G2 para a área de transferência e cola ao invés de escrever
    try:
        pyperclip.copy(campo_G2)  # Copia o valor do campo G2 para a área de transferência
        time.sleep(1)  # Pequena pausa para garantir que a cópia foi concluída
        pyautogui.hotkey('ctrl', 'v')  # Cola o conteúdo usando Ctrl+V
        print(f"  Campo G2 - Descrição detalhada do serviço inserido via cópia/cola.")
    except Exception as e:
        print(f"  Erro ao copiar/colar campo G2: {e}")
        # Fallback: tenta escrever se a cópia falhar
        pyautogui.write(campo_G2, interval=0.1)
        print(f"  Campo G2 inserido via escrita (fallback).")

    time.sleep(1)
    pyautogui.scroll(-800)
    time.sleep(1)
    pyautogui.click(x=1222, y=640)
    print("Clickou no botão avançar")

    #PAGINA 3 VALORES

    time.sleep(2)
    pyautogui.scroll(1000)
    time.sleep(1)
    pyautogui.click(x=216, y=505)
    time.sleep(1)

    # Foco e preenchimento do terceiro campo (Valor)
    try:
        # Usa a string formatada campo_D_formatado
        pyautogui.write(campo_D_formatado, interval=0.1)
        print(f"  Campo D (Valor) '{campo_D_formatado}' inserido.")
    except Exception as e:
        print(f"  Erro ao interagir com terceiro campo (Valor): {e}")
        # pyautogui.hotkey('ctrl', 'w')
        continue

    #Regime especial de tributação 
    time.sleep(1)   
    pyautogui.scroll(-300)
    time.sleep(1)
    pyautogui.click(x=174, y=441)
    time.sleep(1)
    pyautogui.write("nenhum") #microempresa municipal
    time.sleep(1)
    pyautogui.click(x=195, y=480)
    time.sleep(1)

    #Tipo de dedução/redução
    pyautogui.click(x=731, y=441)
    time.sleep(1)
    pyautogui.click(x=731, y=476)
    time.sleep(1)
    pyautogui.click(x=179, y=300)
    time.sleep(1.2)
    
    #ALIGUOTA
    pyautogui.click(x=217, y=640)
    time.sleep(1)
    pyautogui.click(x=559, y=640)
    time.sleep(1)
    pyautogui.press("Backspace", presses=4)
    time.sleep(2)
    pyautogui.write(campo_G4)
    time.sleep(1.5)


    #Situação Tributária PIS/COFINS
    pyautogui.scroll(-400)
    time.sleep(1)
    pyautogui.click(x=177, y=566)
    time.sleep(1)
    pyautogui.write("nen") #nenhum
    time.sleep(1.5)
    pyautogui.click(x=175, y=530)
    time.sleep(1)

    #botão avançar
    pyautogui.scroll(-600)
    time.sleep(1)
    pyautogui.click(x=1202, y=645)
    time.sleep(1)

    #botão EMITIR DPS
    pyautogui.scroll(-1000)
    time.sleep(1)
    pyautogui.click(x=1210, y=640)
    time.sleep(10)

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
messagebox.showinfo("Mensagem do Sistema", "Processo concluído!")
    
    

