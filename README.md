
---

# NFC Automático

## Descrição

Este script automatiza o processo de emissão de NFC através de um formulário web. Ele realiza a leitura de dados a partir de um arquivo Excel, valida e formata os dados, e interage com o navegador para preencher e submeter o formulário. Por fim, o script copia a resposta gerada e a salva de volta no Excel.

## Funcionalidades

- **Leitura de Dados:** Carrega informações de um arquivo Excel utilizando a biblioteca `openpyxl`.
- **Validação e Formatação:** Verifica e formata os campos (como CPF e valor), garantindo que os dados estejam no formato correto para o formulário.
- **Automação Web:** Abre o navegador, preenche os campos do formulário usando `pyautogui`, e interage com a interface (cliques, rolagens e digitação).
- **Extração e Salvamento:** Copia a resposta gerada no formulário e a salva na planilha Excel na coluna especificada.

## Pré-requisitos

- **Python 3.x** instalado no sistema.
- Bibliotecas Python necessárias:
  - [pyautogui](https://pypi.org/project/PyAutoGUI/)
  - [openpyxl](https://pypi.org/project/openpyxl/)
  - [pyperclip](https://pypi.org/project/pyperclip/)
  - As bibliotecas `webbrowser` e `time` já fazem parte da biblioteca padrão do Python.

### Instalação das Dependências

Utilize o `pip` para instalar as bibliotecas necessárias:

```bash
pip install pyautogui openpyxl pyperclip
```

## Configuração

1. **Arquivo Excel:**  
   - Atualize a variável `caminho_arquivo_excel` no script para o caminho correto do seu arquivo Excel.
   - Certifique-se de que a planilha ativa contém os dados a partir da linha definida (por exemplo, a partir da linha 106) e que as colunas estão organizadas da seguinte forma:
     - **Coluna C:** CPF
     - **Coluna D:** Campo B (pode ser uma descrição ou outro dado)
     - **Coluna E:** Valor (que será formatado para duas casas decimais)
     - **Coluna F:** Coluna onde a resposta será salva

2. **Coordenadas de Clique:**  
   - Verifique e, se necessário, ajuste as coordenadas usadas nos comandos `pyautogui.click()` e `pyautogui.moveTo()`. Essas coordenadas devem corresponder à posição dos campos do formulário na sua tela.

3. **URL do Formulário:**  
   - A URL do formulário a ser acessado está definida na variável `url`. Caso o endereço mude, atualize essa variável.

## Uso

Para executar o script, abra o terminal ou prompt de comando e execute:

```bash
python nfc_auto1.py
```

Durante a execução, o script realizará os seguintes passos:
1. Lê os dados do arquivo Excel a partir da linha definida.
2. Para cada linha, valida os dados e verifica se os campos estão corretos.
3. Abre o navegador com a URL especificada e preenche os campos do formulário com os dados lidos.
4. Clica no botão "Emitir" para enviar os dados.
5. Seleciona e copia a resposta do formulário.
6. Salva a resposta copiada na coluna F do Excel.
7. Informa o status de cada etapa no terminal.

## Observações

- **Permissões e Acesso:**  
  Certifique-se de que o arquivo Excel não esteja aberto em outro programa durante a execução, pois isso pode impedir que o script salve as alterações.

- **Ajustes de Tempo:**  
  O script utiliza pausas (`time.sleep()`) para garantir o carregamento das páginas e a execução das interações. Dependendo da velocidade da sua internet e do desempenho do seu computador, pode ser necessário ajustar esses tempos.

- **Resolução da Tela:**  
  As coordenadas definidas para os cliques podem variar de acordo com a resolução e o layout da sua tela. Faça os ajustes necessários para que o script interaja corretamente com os elementos do formulário.

## Troubleshooting

- **Erro ao Carregar o Arquivo Excel:**  
  - Verifique se o caminho definido na variável `caminho_arquivo_excel` está correto.
  - Certifique-se de que o arquivo não está sendo utilizado por outro programa.

- **Coordenadas Incorretas:**  
  - Caso o script não clique no campo correto, ajuste as coordenadas nos comandos `pyautogui.click()` e `pyautogui.moveTo()` conforme a sua necessidade.

- **Erro ao Salvar o Arquivo:**  
  - Se ocorrer um erro de permissão ao salvar o Excel, feche qualquer instância do arquivo aberto em outro programa e verifique as permissões de escrita.

## Contribuições

Contribuições são bem-vindas! Se desejar melhorar este projeto, sinta-se à vontade para abrir issues ou enviar pull requests.

## Licença
