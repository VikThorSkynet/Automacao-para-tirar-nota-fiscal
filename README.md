# NFC Automático

## Visão Geral
Script em Python para preenchimento automático de formulários web a partir de dados do Excel para geração de Nota Fiscal de Consumidor (NFC).

## Requisitos
- Python 3.x
- Bibliotecas: 
  - pyautogui
  - openpyxl
  - time
  - webbrowser
  - win32clipboard

## Configuração
- Atualizar caminho do arquivo Excel no script
- Ajustar `num_linhas` para total de linhas a processar
- Definir coordenadas corretas para interações no formulário web

## Funcionalidade
- Lê dados do Excel (CPF, campos B e C)
- Abre formulário web
- Preenche campos automaticamente
- Clica no botão de envio
- Fecha aba do navegador após processar cada entrada

## Uso
1. Preparar arquivo Excel com dados
2. Ajustar parâmetros do script
3. Executar script
4. Não interromper durante execução

**Observação**: Requer calibração manual das coordenadas de tela para sua resolução específica.
