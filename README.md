# NFC Automação - Sistema de Preenchimento Automático de NFC

## 📋 Descrição
Este projeto automatiza o processo de preenchimento de Notas Fiscais do Consumidor (NFC) através de um script Python que integra dados de uma planilha Excel com um formulário web.

## 🚀 Funcionalidades
- Leitura automática de dados de uma planilha Excel
- Preenchimento automático de formulários web de NFC
- Validação de dados (CPF, campos vazios)
- Captura automática do número da nota fiscal
- Salvamento automático do número da NFC na planilha

## 📦 Pré-requisitos
Para executar este projeto, você precisará ter instalado:

```python
pip install pyautogui
pip install openpyxl
pip install pyperclip
```

## 🛠️ Configuração
1. Estrutura da planilha Excel necessária:
   - Coluna C: CPF
   - Coluna D: Campo de preenchimento 1
   - Coluna E: Campo de preenchimento 2
   - Coluna F: Número da NFC (preenchido automaticamente)

2. Ajuste as configurações no código:
   - Caminho do arquivo Excel
   - Linha inicial e número de linhas a serem processadas
   - URL do formulário web

## 💻 Como Usar
1. Prepare sua planilha Excel com os dados necessários
2. Ajuste as variáveis `linha_inicial` e `num_linhas` no código
3. Execute o script:
```python
python nfc_auto.py
```

## ⚠️ Considerações Importantes
- O script utiliza coordenadas de tela específicas (pyautogui). Pode ser necessário ajustar as coordenadas de acordo com sua resolução de tela
- Mantenha o arquivo Excel fechado durante a execução do script
- O script inclui delays (time.sleep) para garantir o carregamento adequado das páginas
- Certifique-se de ter uma conexão estável com a internet

## 🔍 Validações
O script inclui as seguintes validações:
- Verifica CPFs vazios
- Verifica campos vazios ou marcados com "-"
- Verifica células com destaque em amarelo
- Tratamento de erros ao salvar no Excel

## 🚫 Tratamento de Erros
- Verifica permissão de escrita no arquivo Excel
- Trata erros na leitura da área de transferência
- Validação de campos obrigatórios

## ⚙️ Personalização
Para ajustar as coordenadas de clique:
1. Use `pyautogui.position()` em um console Python separado
2. Mova o mouse para a posição desejada
3. Anote as coordenadas x,y
4. Atualize no código as coordenadas nos comandos `pyautogui.click()`

## 📝 Logs
O script fornece feedback através do console sobre:
- Linhas sendo processadas
- Erros encontrados
- Status do processo

## 🤝 Contribuição
Sinta-se à vontade para contribuir com o projeto através de:
- Relatórios de bugs
- Sugestões de melhorias
- Pull requests

## 📄 Licença
Este projeto está sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

## 🎯 Dicas de Uso
1. Mantenha o navegador como janela ativa durante a execução
2. Não mova o mouse durante a execução do script
3. Verifique periodicamente o arquivo Excel para garantir o correto salvamento
4. Faça backup dos dados antes de executar o script

## 🔧 Solução de Problemas
Se encontrar problemas:
1. Verifique se todos os módulos estão instalados
2. Confirme se as coordenadas de clique estão corretas para sua tela
3. Ajuste os tempos de espera (time.sleep) se necessário
4. Verifique se o arquivo Excel está acessível e não está aberto
