from flask import Flask, request, render_template_string, send_file
import pandas as pd
import smtplib
import ssl
import random
import time
import tempfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
import os

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Envío Masivo SMTP</title>
</head>
<body>
    <h1>Envío de Correos Masivos SMTP</h1>
    <form method="POST" enctype="multipart/form-data">
        <label>Archivo de Cuentas SMTP (Excel):</label><br>
        <input type="file" name="archivo_cuentas" required><br><br>
        
        <label>Archivo de Destinatarios (Excel):</label><br>
        <input type="file" name="archivo_destinatarios" required><br><br>

        <label>Archivo HTML del Cuerpo del Correo:</label><br>
        <input type="file" name="archivo_html" required><br><br>

        <label>Archivo de Asuntos (Excel o TXT):</label><br>
        <input type="file" name="archivo_asuntos" required><br><br>

        <label>Cantidad máxima de correos por cuenta:</label><br>
        <input type="number" name="max_correos" value="30" min="1" required><br><br>

        <button type="submit">Iniciar Envío</button>
    </form>

    {% if resultados %}
        <h2>Resultados del Envío:</h2>
        <ul>
        {% for resultado in resultados %}
            <li>{{ resultado }}</li>
        {% endfor %}
        </ul>
    {% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    resultados = []
    detalles_envios = []
    if request.method == 'POST':
        archivo_cuentas = request.files['archivo_cuentas']
        archivo_destinatarios = request.files['archivo_destinatarios']
        archivo_html = request.files['archivo_html']
        archivo_asuntos = request.files['archivo_asuntos']
        max_correos = int(request.form['max_correos'])

        cuentas = pd.read_excel(archivo_cuentas)
        destinatarios = pd.read_excel(archivo_destinatarios)
        cuerpo_html = archivo_html.read().decode('utf-8')

        if archivo_asuntos.filename.endswith('.xlsx'):
            asuntos_df = pd.read_excel(archivo_asuntos)
            lista_asuntos = asuntos_df.iloc[:, 0].dropna().tolist()
        else:
            lista_asuntos = archivo_asuntos.read().decode('utf-8').splitlines()
            lista_asuntos = [line.strip() for line in lista_asuntos if line.strip()]

        contexto = ssl.create_default_context()

        ultimo_email_remitente = None
        ultimo_password = None
        ultimo_smtp_server = None
        ultimo_smtp_port = None
        ultimo_nombre_remitente = None

        for idx, cuenta in cuentas.iterrows():
            email_remitente = cuenta['Email']
            password = cuenta['Password']
            smtp_server = cuenta['SMTP Server']
            smtp_port = cuenta['SMTP Port']
            nombre_remitente = cuenta.get('Nombre Remitente', email_remitente)

            ultimo_email_remitente = email_remitente
            ultimo_password = password
            ultimo_smtp_server = smtp_server
            ultimo_smtp_port = smtp_port
            ultimo_nombre_remitente = nombre_remitente

            try:
                with smtplib.SMTP_SSL(smtp_server, smtp_port, context=contexto) as server:
                    server.login(email_remitente, password)
                    destinatarios_sample = destinatarios.sample(n=min(max_correos, len(destinatarios)))

                    for _, row in destinatarios_sample.iterrows():
                        mensaje = MIMEMultipart("alternative")
                        mensaje["Subject"] = random.choice(lista_asuntos)
                        mensaje["From"] = formataddr((nombre_remitente, email_remitente))
                        mensaje["To"] = row['mail']

                        parte_html = MIMEText(cuerpo_html, "html")
                        mensaje.attach(parte_html)

                        try:
                            server.sendmail(email_remitente, row['mail'], mensaje.as_string())
                            resultados.append(f"✅ Enviado a {row['mail']} desde {nombre_remitente} <{email_remitente}>")
                            detalles_envios.append({"Emisor": email_remitente, "Destinatario": row['mail'], "Estado": "Enviado"})
                        except Exception as e:
                            resultados.append(f"❌ Error enviando a {row['mail']} desde {nombre_remitente} <{email_remitente}>: {str(e)}")
                            detalles_envios.append({"Emisor": email_remitente, "Destinatario": row['mail'], "Estado": f"Error: {str(e)}"})

                        time.sleep(random.randint(2, 5))

                espera_cuenta = random.randint(2, 7) * 60
                time.sleep(espera_cuenta)

            except Exception as e:
                resultados.append(f"❌ Error conectando con {nombre_remitente} <{email_remitente}>: {str(e)}")

        # Crear el reporte Excel
        if detalles_envios:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                reporte_path = tmp.name

            writer = pd.ExcelWriter(reporte_path, engine='openpyxl')
            df_detalles = pd.DataFrame(detalles_envios)
            df_detalles.to_excel(writer, sheet_name='DetalleEnvios', index=False)
            resumen = df_detalles.groupby('Emisor').size().reset_index(name='Cantidad de Correos Enviados')
            resumen.to_excel(writer, sheet_name='ResumenEmisores', index=False)
            writer.save()

            # Enviar el reporte al correo deseado
            try:
                mensaje_reporte = MIMEMultipart()
                mensaje_reporte["Subject"] = "Reporte de Envíos Masivos"
                mensaje_reporte["From"] = formataddr((ultimo_nombre_remitente, ultimo_email_remitente))
                mensaje_reporte["To"] = "cuevasarancibia50@gmail.com"

                cuerpo_texto = MIMEText("Adjunto reporte de los envíos realizados.", "plain")
                mensaje_reporte.attach(cuerpo_texto)

                with open(reporte_path, "rb") as adjunto:
                    parte_adjunto = MIMEBase("application", "octet-stream")
                    parte_adjunto.set_payload(adjunto.read())
                    encoders.encode_base64(parte_adjunto)
                    parte_adjunto.add_header("Content-Disposition", f"attachment; filename=reporte_envios.xlsx")
                    mensaje_reporte.attach(parte_adjunto)

                with smtplib.SMTP_SSL(ultimo_smtp_server, ultimo_smtp_port, context=contexto) as server:
                    server.login(ultimo_email_remitente, ultimo_password)
                    server.sendmail(ultimo_email_remitente, "cuevasarancibia50@gmail.com", mensaje_reporte.as_string())

            except Exception as e:
                resultados.append(f"❌ Error enviando reporte final: {str(e)}")

    return render_template_string(HTML_TEMPLATE, resultados=resultados)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

