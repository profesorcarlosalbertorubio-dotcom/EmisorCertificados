import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert

def obtener_fecha_dinamica(idioma="ESP"):
    """
    Genera la fecha dinámica del sistema con el formato requerido.
    Para Español (ESP): Nov/01/2026
    Para Inglés (ENG): Nov/01/2026 (meses abreviados en inglés)
    """
    ahora = datetime.now()
    dia = str(ahora.day).zfill(2)
    anio = str(ahora.year)
    
    if idioma.upper() == "ENG":
        meses = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    else:
        meses = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
        
    mes_3 = meses.get(ahora.month, "Nov")
    return f"{mes_3}/{dia}/{anio}"

def generar_certificados_bilingues(nombre_alumno):
    """
    Genera en simultáneo el certificado en Español y el certificado en Inglés
    utilizando BorraCert-ESP.docx y BorraCert-ENG.docx con fecha dinámica.
    """
    os.makedirs("salida", exist_ok=True)
    nombre_limpio = nombre_alumno.replace(" ", "_")
    
    print(f"\n--- INICIANDO EMISIÓN BILINGÜE PARA: {nombre_alumno} ---")

    # 1. GENERAR CERTIFICADO EN ESPAÑOL
    path_es = "Insumos/BorraCert-ESP.docx"
    temp_es = f"salida/temp_{nombre_limpio}_ESP.docx"
    pdf_es = f"salida/Certificado_{nombre_limpio}_ESP.pdf"
    
    try:
        doc_es = DocxTemplate(path_es)
        fecha_es = obtener_fecha_dinamica("ESP")
        doc_es.render({'nombre': nombre_alumno, 'fecha': fecha_es})
        doc_es.save(temp_es)
        convert(temp_es, pdf_es)
        print(f"[ÉXITO] Certificado Español generado: {pdf_es}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el certificado en Español: {e}")
    finally:
        if os.path.exists(temp_es):
            os.remove(temp_es)

    # 2. GENERAR CERTIFICADO EN INGLÉS
    path_en = "Insumos/BorraCert-ENG.docx"
    temp_en = f"salida/temp_{nombre_limpio}_ENG.docx"
    pdf_en = f"salida/Certificado_{nombre_limpio}_ENG.pdf"
    
    try:
        doc_en = DocxTemplate(path_en)
        fecha_en = obtener_fecha_dinamica("ENG")
        doc_en.render({'nombre': nombre_alumno, 'fecha': fecha_en})
        doc_en.save(temp_en)
        convert(temp_en, pdf_en)
        print(f"[ÉXITO] Certificado Inglés generado: {pdf_en}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el certificado en Inglés: {e}")
    finally:
        if os.path.exists(temp_en):
            os.remove(temp_en)

    print("--- PROCESO BILINGÜE FINALIZADO ---\n")

if __name__ == "__main__":
    # Prueba local con su nombre
    generar_certificados_bilingues("Carlos Alberto Rubio")