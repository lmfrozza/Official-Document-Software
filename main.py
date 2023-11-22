import pandas as pd
from docx import Document
import random  # Para gerar dados aleatórios (nomes, endereços, etc.)
from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout

GUI = Builder.load_file("tela.kv")
arquivo_xlsx= 'planilhamodel.xlsx'
arquivo_docx = 'ofcmodelo.docs'
# Lê os dados da planilha usando o Pandas

class ODsoftware(App):
    def build(self):
        return GUI
    
    def on_start(self):
        # Obter o botão do arquivo .kv
        button = self.root.ids.button

        # Atribuir a função ao botão usando a propriedade 'on_press'
        button.bind(on_press=self.iniciate)
    def iniciate(self, instance):
        arquivo_xlsx = self.root.ids.xlsx
        arquivo_docx = self.root.ids.docx
        
        planilha = pd.read_excel(arquivo_xlsx)  # Substitua "sua_planilha.xlsx" pelo caminho da sua planilha
        for index, row in planilha.iterrows():
            C1 = row["C1"]  # Substitua "Nome" pelo nome da coluna na planilha
            C2 = row["C2"]  # Substitua "Endereço" pelo nome da coluna na planilha

            # Abre um novo modelo de documento para cada iteração
            documento_modelo = Document(arquivo_docx)  # Substitua "modelo.docx" pelo caminho do seu modelo de documento

            # Aqui você pode adicionar mais campos conforme necessário

            # Preenche o documento modelo com os dados
            for paragraph in documento_modelo.paragraphs:
                if "{C1}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C1}", str(C1))
                if "{C2}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C2}", C2)
            # Salva o documento com um nome único, por exemplo, com base no nome da pessoa
            documento_modelo.save(f"Ofício Nº{C1} - {C2}.docx")



ODsoftware().run()