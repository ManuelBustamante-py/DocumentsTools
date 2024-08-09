from comtypes import client

def word_to_pdf(input_docx, output_pdf):
    word = client.CreateObject('Word.Application')
    word.Visible = False  # word se oculta luego de abrirse y ser reconocido por el sistema
    doc = word.Documents.Open(input_docx)
    doc.SaveAs(output_pdf, FileFormat=17)  # asigné formato de archivo correspondiente a pdf
    doc.Close()

    word.Quit()

if __name__ == "__main__":
    # Ruta completa al archivo DOCX de entrada, o el que se desee convertir a pdf
    input_docx = r" .docx"
    
    # Nombre y ruta para el archivo PDF de salida, donde se quiere guardar el archivo
    output_pdf = r" .pdf"
    
    word_to_pdf(input_docx, output_pdf)
    print(f'Archivo convertido: {output_pdf}')