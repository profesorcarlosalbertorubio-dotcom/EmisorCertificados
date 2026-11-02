import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert

def obtener_fecha_dinamica_estilo():
    """
    Obtiene la fecha actual del sistema de forma dinámica y la formatea 
    exactamente como: Nov/01/2026
    """
    meses_cortos = {
        1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr",
        5: "May", 6: "Jun", 7: "Jul", 8: "Ago",
        9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
    }
    
    ahora = datetime.now()
    dia_dos_digitos = str(ahora.day).zfill(2)  # Ej. "01", "15"
    mes_num = ahora.month
    anio = str(ahora.year)
    
    mes_3 = meses_cortos.get(mes_num, "Nov")
    
    # Formato final exacto requerido: Nov/01/2026
    return f"{mes_3}/{dia_dos_digitos}/{anio}"

def generar_certificado_desde_word(nombre_alumno):
    """
    Toma la plantilla visual de Word, inyecta el nombre y la fecha 
    actual del sistema de forma dinámica, y exporta a PDF limpiamente.
    """
    path_plantilla_word = "Insumos/BorraCert.docx"
    os.makedirs("salida", exist_ok=True)
    
    print(f"Generando certificado para: {nombre_alumno}...")

    # 1. Cargar la plantilla visual de Word
    try:
        doc = DocxTemplate(path_plantilla_word)
    except FileNotFoundError:
        print(f"Error: No se encontró {path_plantilla_word} en la carpeta Insumos.")
        return

    # 2. Generar la fecha del día actual de forma dinámica
    fecha_actual_formateada = obtener_fecha_dinamica_estilo()
    print(f"Fecha detectada del sistema: {fecha_actual_formateada}")

    # 3. Contexto unificado para reemplazar en el documento Word
    contexto = {
        'nombre': nombre_alumno,
        'fecha': fecha_actual_formateada
    }
    
    doc.render(contexto)

    # 4. Guardar archivo temporal y convertir a PDF en la carpeta salida/
    nombre_limpio = nombre_alumno.replace(" ", "_")
    docx_temporal = f"salida/temp_{nombre_limpio}.docx"
    pdf_final = f"salida/Certificado_{nombre_limpio}_ES.pdf"

    doc.save(docx_temporal)
    
    # Convierte el archivo de Word a PDF utilizando Word de forma nativa
    try:
        convert(docx_temporal, pdf_final)
        print(f"¡Éxito total! PDF generado correctamente en: {pdf_final}\n")
    except Exception as e:
        print(f"Error al convertir a PDF: {e}")
    finally:
        # Limpieza del archivo temporal
        if os.path.exists(docx_temporal):
            os.remove(docx_temporal)

if __name__ == "__main__":
    print("--- AUTOMATIZACIÓN DINÁMICA DE CERTIFICADOS (MUNDO LABORAL & TECH) ---")
    generar_certificado_desde_word("Carlos Alberto Rubio")