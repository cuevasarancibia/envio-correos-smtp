from flask import Flask, request, render_template_string
import pandas as pd
import smtplib
import ssl
import random
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
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
    if request.method == 'POST':
        archivo_cuentas = request.files['archivo_cuentas']
        archivo_destinatarios = request.files['archivo_destinatarios']
        archivo_html = request.files['archivo_html']
        archivo_asuntos = request.files['archivo_asuntos']
        max_correos = int(request.form['max_correos'])

        cuentas = pd.read_excel(archivo_cuentas)
        destinatarios = pd.read_excel(archivo_destinatarios)
        cuerpo_html = archivo_html.read().decode('utf-8')

        # Cargar asuntos desde Excel o TXT
        if archivo_asuntos.filename.endswith('.xlsx'):
            asuntos_df = pd.read_excel(archivo_asuntos)
            lista_asuntos = asuntos_df.iloc[:, 0].dropna().tolist()
        else:
            lista_asuntos = archivo_asuntos.read().decode('utf-8').splitlines()
            lista_asuntos = [line.strip() for line in lista_asuntos if line.strip()]

        contexto = ssl.create_default_context()

        for idx, cuenta in cuentas.iterrows():
            email_remitente = cuenta['Email']
            password = cuenta['Password']
            smtp_server = cuenta['SMTP Server']
            smtp_port = cuenta['SMTP Port']
            nombre_remitente = cuenta.get('Nombre Remitente', email_remitente)

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
                        except Exception as e:
                            resultados.append(f"❌ Error enviando a {row['mail']} desde {nombre_remitente} <{email_remitente}>: {str(e)}")

                        time.sleep(random.randint(2, 5))

                espera_cuenta = random.randint(2, 7) * 60
                time.sleep(espera_cuenta)

            except Exception as e:
                resultados.append(f"❌ Error conectando con {nombre_remitente} <{email_remitente}>: {str(e)}")

    return render_template_string(HTML_TEMPLATE, resultados=resultados)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
