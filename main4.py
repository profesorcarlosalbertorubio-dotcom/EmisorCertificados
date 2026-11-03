import os
import subprocess
from datetime import datetime
from flask import Flask, render_template, request, jsonify
from docxtpl import DocxTemplate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

# ================= CONFIGURACIÓN DE CORREO (MundoLaboralTech) =================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
CORREO_EMISOR = "profesor.carlosalbertorubio@gmail.com"
CONTRASENA_EMISOR = "pqjn xkcp ovvv uyjn"  # Sus 16 caracteres sin espacios

def convertir_docx_a_pdf(input_path, output_dir):
    """
    Convierte un archivo .docx a PDF utilizando LibreOffice en entorno Linux (Render).
    """
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", 
        input_path, "--outdir", output_dir
    ], check=True)

def obtener_fecha_dinamica(idioma="ESP"):
    """
    Genera la fecha dinámica del sistema con el formato requerido.
    Para Español (ESP): Nov/02/2026
    Para Inglés (ENG): Nov/02/2026 (meses abreviados en inglés)
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

def enviar_correo_certificados(nombre_alumno, correo_destino, pdf_es, pdf_en):
    """
    Envía por SMTP seguro de Gmail los certificados adjuntos (Español e Inglés).
    """
    print(f"\n[CORREO] Preparando envío para {nombre_alumno} ({correo_destino})...")
    
    msg = MIMEMultipart()
    msg['From'] = CORREO_EMISOR
    msg['To'] = correo_destino
    msg['Subject'] = f"Sus Certificados Bilingües - @MundoLaboralTech"

    cuerpo = f"""Estimado/a {nombre_alumno},

Le saludamos desde @MundoLaboralTech. Adjunto a este mensaje encontrará sus certificados oficiales en Español y en Inglés correspondientes a su proceso de formación.

Atentamente,
Profesor Carlos Alberto Rubio Sánchez
@MundoLaboralTech
"""
    msg.attach(MIMEText(cuerpo, 'plain'))

    # Adjuntar PDF en Español
    try:
        with open(pdf_es, "rb") as adjunto_es:
            parte_es = MIMEBase('application', 'octet-stream')
            parte_es.set_payload(adjunto_es.read())
        encoders.encode_base64(parte_es)
        parte_es.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(pdf_es)}")
        msg.attach(parte_es)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo adjuntar el PDF en español: {e}")

    # Adjuntar PDF en Inglés
    try:
        with open(pdf_en, "rb") as adjunto_en:
            parte_en = MIMEBase('application', 'octet-stream')
            parte_en.set_payload(adjunto_en.read())
        encoders.encode_base64(parte_en)
        parte_en.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(pdf_en)}")
        msg.attach(parte_en)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo adjuntar el PDF en inglés: {e}")

    # Conexión SMTP y Envío
    try:
        servidor = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        servidor.login(CORREO_EMISOR, CONTRASENA_EMISOR)
        servidor.sendmail(CORREO_EMISOR, correo_destino, msg.as_string())
        servidor.quit()
        print(f"[ÉXITO] ¡Correo enviado satisfactoriamente a {correo_destino}!")
        return True
    except Exception as e:
        print(f"[ERROR DE ENVÍO] No se pudo enviar el correo: {e}")
        return False

def generar_certificados_bilingues(nombre_alumno, correo_destino):
    """
    Genera en simultáneo el certificado en Español y el certificado en Inglés.
    """
    os.makedirs("salida", exist_ok=True)
    nombre_limpio = nombre_alumno.replace(" ", "_")
    
    print(f"\n--- INICIANDO EMISIÓN BILINGÜE PARA: {nombre_alumno} ---")

    # 1. CERTIFICADO ESPAÑOL
    path_es = "Insumos/BorraCert-ESP.docx"
    temp_es = f"salida/temp_{nombre_limpio}_ESP.docx"
    pdf_es = f"salida/Certificado_{nombre_limpio}_ESP.pdf"
    
    try:
        doc_es = DocxTemplate(path_es)
        fecha_es = obtener_fecha_dinamica("ESP")
        doc_es.render({'nombre': nombre_alumno, 'fecha': fecha_es})
        doc_es.save(temp_es)
        convertir_docx_a_pdf(temp_es, "salida")
        print(f"[ÉXITO] Certificado Español generado: {pdf_es}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el certificado en Español: {e}")
    finally:
        if os.path.exists(temp_es):
            os.remove(temp_es)

    # 2. CERTIFICADO INGLÉS
    path_en = "Insumos/BorraCert-ENG.docx"
    temp_en = f"salida/temp_{nombre_limpio}_ENG.docx"
    pdf_en = f"salida/Certificado_{nombre_limpio}_ENG.pdf"
    
    try:
        doc_en = DocxTemplate(path_en)
        fecha_en = obtener_fecha_dinamica("ENG")
        doc_en.render({'nombre': nombre_alumno, 'fecha': fecha_en})
        doc_en.save(temp_en)
        convertir_docx_a_pdf(temp_en, "salida")
        print(f"[ÉXITO] Certificado Inglés generado: {pdf_en}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el certificado en Inglés: {e}")
    finally:
        if os.path.exists(temp_en):
            os.remove(temp_en)

    # 3. ENVÍO POR CORREO
    exito = False
    if os.path.exists(pdf_es) and os.path.exists(pdf_en):
        exito = enviar_correo_certificados(nombre_alumno, correo_destino, pdf_es, pdf_en)
    else:
        print("[ERROR] Faltan archivos PDF para realizar el envío por correo.")

    print("--- PROCESO BILINGÜE FINALIZADO ---\n")
    return exito

@app.route("/")
def index():
    return "Emisor de Certificados Bilingües @MundoLaboralTech funcionando correctamente."

@app.route("/procesar", methods=["POST"])
def procesar():
    """
    Ruta receptora invocada automáticamente desde Google Sheets.
    """
    datos = request.get_json(silent=True) or {}
    nombre = datos.get("nombre")
    correo = datos.get("correo")

    if not nombre or not correo:
        return jsonify({"status": "error", "mensaje": "Datos incompletos"}), 400

    exito = generar_certificados_bilingues(nombre, correo)
    if exito:
        return jsonify({"status": "ok", "mensaje": f"Certificados enviados a {correo}"}), 200
    else:
        return jsonify({"status": "error", "mensaje": "Fallo la generación o el envío"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)