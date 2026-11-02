import os
from docxtpl import DocxTemplate
from docx2pdf import convert

def formatear_mes_tres_letras(mes_str):
    """Extrae las tres primeras letras del mes (ej. Noviembre -> Nov)."""
    meses_cortos = {
        "enero": "Ene", "febrero": "Feb", "marzo": "Mar", "abril": "Abr",
        "mayo": "May", "junio": "Jun", "julio": "Jul", "agosto": "Ago",
        "septiembre": "Sep", "octubre": "Oct", "noviembre": "Nov", "diciembre": "Dic"
    }
    mes_limpio = mes_str.strip().lower()
    return meses_cortos.get(mes_limpio, mes_str[:3].capitalize())

def generar_certificado_desde_word(nombre_alumno, dia="15", mes="Noviembre"):
    """
    Toma la plantilla visual de Word (BorraCert.docx), rellena los campos 
    y la convierte a PDF sin alterar el diseño visual original.
    """
    path_plantilla_word = "Insumos/BorraCert.docx"
    os.makedirs("salida", exist_ok=True)
    
    print(f"Generando certificado desde plantilla Word para: {nombre_alumno}...")

    # 1. Cargar la plantilla visual de Word
    try:
        doc = DocxTemplate(path_plantilla_word)
    except FileNotFoundError:
        print(f"Error: No se encontró {path_plantilla_word} en la carpeta Insumos.")
        return

    # 2. Reemplazar los datos (respetando formato, tamaño y color de Word)
    mes_3 = formatear_mes_tres_letras(mes)
    contexto = {
        'nombre': nombre_alumno,
        'dia': dia,
        'mes': mes_3
    }
    
    doc.render(contexto)

    # 3. Guardar borrador en Word y convertir a PDF en la carpeta salida/
    nombre_limpio = nombre_alumno.replace(" ", "_")
    docx_temporal = f"salida/temp_{nombre_limpio}.docx"
    pdf_final = f"salida/Certificado_{nombre_limpio}_ES.pdf"

    doc.save(docx_temporal)
    
    # Convierte el archivo de Word a PDF utilizando Word de forma nativa
    try:
        convert(docx_temporal, pdf_final)
        print(f"¡Éxito total! PDF generado en: {pdf_final}\n")
    except Exception as e:
        print(f"Error al convertir a PDF: {e}")
    finally:
        # Limpieza del archivo temporal
        if os.path.exists(docx_temporal):
            os.remove(docx_temporal)

if __name__ == "__main__":
    print("--- AUTOMATIZACIÓN WORD A PDF (MUNDO LABORAL & TECH) ---")
    generar_certificado_desde_word("Carlos Alberto Rubio", "15", "Noviembre")