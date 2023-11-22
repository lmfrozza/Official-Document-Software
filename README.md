# README

## Overview
Este é um script em Python que utiliza as bibliotecas pandas, docx e kivy para automatizar a geração de documentos do tipo ofício a partir de dados contidos em uma planilha Excel (.xlsx). O script lê os dados da planilha, preenche um modelo de documento (.docx) com esses dados e salva cada documento gerado com um nome único, baseado em uma variável numérica.

## Pré-requisitos
Certifique-se de ter as seguintes bibliotecas instaladas no seu ambiente Python antes de executar o script:
- pandas
- docx
- kivy

Você pode instalá-las usando o seguinte comando:
```bash
pip install pandas docx kivy
```
# Utilização
Certifique-se de ter uma planilha Excel com os dados que você deseja processar. O caminho da planilha deve ser fornecido no arquivo KV (tela.kv).
Tenha um modelo de documento (.docx) que contém marcadores de posição para os dados da planilha. O caminho do modelo deve ser fornecido no arquivo KV (tela.kv).
Execute o script e utilize a interface gráfica fornecida para iniciar o processo de geração de documentos.

# Estrutura do Código
O código principal está contido no arquivo Python (script.py).
O arquivo KV (tela.kv) contém a descrição da interface gráfica usando a linguagem Kivy.
A lógica de geração de documentos está na classe ODsoftware que herda da classe App do Kivy.

# Personalização
Personalize os marcadores de posição no modelo do documento (.docx) conforme necessário, adicionando ou removendo campos no loop de preenchimento.

# Observações
Certifique-se de substituir os placeholders como "{C1}", "{C2}", etc., no seu modelo de documento pelo nome das colunas reais na sua planilha Excel.
Os documentos gerados serão salvos numericamente com base na variável cont, garantindo nomes únicos.
Espero que estas informações sejam úteis! Se precisar de mais ajuda ou tiver alguma dúvida, estou à disposição.
