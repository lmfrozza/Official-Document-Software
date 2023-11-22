import pandas as pd
from docx import Document
import random  # Para gerar dados aleatórios (nomes, endereços, etc.)
from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout

GUI = Builder.load_file("tela.kv")
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
        
        
        planilha = pd.read_excel(f'{arquivo_xlsx.text}')  # Substitua "sua_planilha.xlsx" pelo caminho da sua planilha
        for index, row in planilha.iterrows():
            C1 = row["C1"]  
            C2 = row["C2"]  
            C3 = row["C3"]  
            C4 = row["C4"]
            C5 = row["C5"] 
            C6 = row["C6"]
            # Abre um novo modelo de documento para cada iteração
            documento_modelo = Document(f'{arquivo_docx.text}')  # Substitua "modelo.docx" pelo caminho do seu modelo de documento

            # Aqui você pode adicionar mais campos conforme necessário

            # Preenche o documento modelo com os dados
            for paragraph in documento_modelo.paragraphs:
                if "{C1}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C1}", str(C1))
                if "{C2}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C2}", str(C2))
                if "{C3}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C3}", str(C3))
                if "{C4}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C4}", str(C4))
                if "{C5}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C5}", str(C5))
                if "{C6}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{C6}", str(C6))
            # Salva o documento com um nome único, por exemplo, com base no nome da pessoa
            documento_modelo.save(f"Ofício Nº{C1} - {C2}.docx")



ODsoftware().run()