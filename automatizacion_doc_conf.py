import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

doc = DocxTemplate('/Users/juanpablomoya/Desktop/proyecto/conf_fact.docx')

nombre = 'Juan Pablo Moya'
correo = 'jpmoya@tsgroup.cl'
fecha = datetime.today().strftime('%d/%m/%y')


datos = {'nombre':nombre, 'correo':correo, 'fecha':fecha}

df = pd.read_excel('/Users/juanpablomoya/Desktop/proyecto/docs_noconfirmados.xlsx')


for indice, fila in df.iterrows():
    fullemail = {
        'destinatario':fila['Responsable'],
        'area':fila['proveedor'],
        'doc1': fila['doc1'],
        'doc2': fila['doc2'],
        'doc3': fila['doc3'],
        'doc4': fila['doc4']
                 }
    fullemail.update(datos)
    print(fullemail)

    doc.render(fullemail)
    doc.save(f'Email_for_{fila["Responsable"]}.docx')










