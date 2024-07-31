import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

#Subimos la plantilla que tiene el cuerpo de nuestro correo en word.
doc = DocxTemplate('conf_fact.docx')

#Agregamos las constantes que tendrá nuestro correo.
nombre = 'Juan Pablo Moya'
correo = 'jpmoya@tsgroup.cl'
fecha = datetime.today().strftime('%d/%m/%y')


datos = {'nombre':nombre, 'correo':correo, 'fecha':fecha}

#importamos la base de datos donde están nuestros documentos pendientes de confirmación.
df = pd.read_excel('/Users/juanpablomoya/Desktop/proyecto/docs_noconfirmados.xlsx')


#Este bucle for itera por cada dato que tiene nuestra bd y los agrega a nuestro word con las constantes.
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

    
#Renderizamos el email completo y guardamos por cada uno de los responsables, queda listo para enviar y editar según corresponda.
    doc.render(fullemail)
    doc.save(f'Email_for_{fila["Responsable"]}.docx')










