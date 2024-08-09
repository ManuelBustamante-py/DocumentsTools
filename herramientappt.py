from comtypes import client

def ppt_to_pdf(input_ppt, output_pdf):
    powerpoint = client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1  # en este caso nos aseguramos de que powerpoint sea visible
    ppt = powerpoint.Presentations.Open(input_ppt)
    ppt.SaveAs(output_pdf, 32)  # para esta situación, el formato de archivo es 32, correspondeinte a pdf
    ppt.Close()
    powerpoint.Quit()

if __name__ == "__main__":
    # Ruta completa al archivo PPT de entrada, el archivo que se desea convertir a pdf
    input_ppt = r" .pptx"
    
    # Nombre y ruta para el archivo PDF de salida, el lugar donde se quiere guardar el archivo
    output_pdf = r" .pdf"
    
    ppt_to_pdf(input_ppt, output_pdf)
    print(f'Archivo convertido: {output_pdf}')