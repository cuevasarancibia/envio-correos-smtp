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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Envío Masivo SMTP - Los 100</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        :root {
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --success: #10b981;
            --error: #ef4444;
            --bg: #f8fafc;
            --bg-card: #ffffff;
            --text: #1e293b;
            --text-light: #64748b;
            --border: #e2e8f0;
            --shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
            --shadow-lg: 0 10px 15px rgba(0, 0, 0, 0.1);
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, var(--bg) 0%, #f1f5f9 100%);
            color: var(--text);
            min-height: 100vh;
            padding: 2rem;
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
        }

        header {
            text-align: center;
            margin-bottom: 3rem;
            animation: slideDown 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        h1 {
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
            background: linear-gradient(135deg, var(--primary) 0%, #1e40af 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        .subtitle {
            color: var(--text-light);
            font-size: 1.1rem;
        }

        .form-card {
            background: var(--bg-card);
            border-radius: 12px;
            padding: 2rem;
            box-shadow: var(--shadow);
            margin-bottom: 2rem;
            animation: fadeInUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1) 0.1s both;
        }

        .form-group {
            margin-bottom: 1.5rem;
            animation: fadeInUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        .form-group:nth-child(1) { animation-delay: 0.2s; }
        .form-group:nth-child(2) { animation-delay: 0.3s; }
        .form-group:nth-child(3) { animation-delay: 0.4s; }
        .form-group:nth-child(4) { animation-delay: 0.5s; }
        .form-group:nth-child(5) { animation-delay: 0.6s; }

        label {
            display: block;
            font-weight: 600;
            margin-bottom: 0.5rem;
            color: var(--text);
            font-size: 0.95rem;
            transition: color 0.3s ease;
        }

        input[type="file"],
        input[type="number"] {
            width: 100%;
            padding: 0.75rem;
            border: 2px solid var(--border);
            border-radius: 8px;
            font-size: 1rem;
            transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
            background: var(--bg);
        }

        input[type="file"]:hover,
        input[type="number"]:hover {
            border-color: var(--primary);
            background: #f0f4ff;
        }

        input[type="file"]:focus,
        input[type="number"]:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
            transform: translateY(-2px);
        }

        button {
            width: 100%;
            padding: 1rem;
            background: linear-gradient(135deg, var(--primary) 0%, #1e40af 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 1.1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
            box-shadow: var(--shadow);
            animation: fadeInUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1) 0.7s both;
        }

        button:hover {
            background: linear-gradient(135deg, var(--primary-hover) 0%, #1e3a8a 100%);
            transform: translateY(-3px);
            box-shadow: var(--shadow-lg);
        }

        button:active {
            transform: translateY(-1px);
        }

        .results-card {
            background: var(--bg-card);
            border-radius: 12px;
            padding: 2rem;
            box-shadow: var(--shadow);
            animation: slideUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        .results-card h2 {
            font-size: 1.5rem;
            margin-bottom: 1.5rem;
            color: var(--text);
            animation: fadeIn 0.6s ease-out 0.2s both;
        }

        .results-list {
            list-style: none;
        }

        .results-list li {
            padding: 1rem;
            margin-bottom: 0.75rem;
            background: var(--bg);
            border-radius: 8px;
            border-left: 4px solid var(--primary);
            animation: slideInLeft 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
            transition: all 0.3s ease;
        }

        .results-list li:hover {
            transform: translateX(4px);
            box-shadow: var(--shadow);
        }

        .results-list li:nth-child(1) { animation-delay: 0.1s; }
        .results-list li:nth-child(2) { animation-delay: 0.2s; }
        .results-list li:nth-child(3) { animation-delay: 0.3s; }
        .results-list li:nth-child(4) { animation-delay: 0.4s; }
        .results-list li:nth-child(5) { animation-delay: 0.5s; }

        .results-list li:has(✅) {
            border-left-color: var(--success);
            background: rgba(16, 185, 129, 0.05);
        }

        .results-list li:has(❌) {
            border-left-color: var(--error);
            background: rgba(239, 68, 68, 0.05);
        }

        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @keyframes slideUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @keyframes fadeIn {
            from {
                opacity: 0;
            }
            to {
                opacity: 1;
            }
        }

        @keyframes slideInLeft {
            from {
                opacity: 0;
                transform: translateX(-20px);
            }
            to {
                opacity: 1;
                transform: translateX(0);
            }
        }

        @media (max-width: 768px) {
            body {
                padding: 1rem;
            }

            h1 {
                font-size: 1.75rem;
            }

            .form-card {
                padding: 1.5rem;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📧 Envío Masivo SMTP</h1>
            <p class="subtitle">Los 100 - Sistema profesional de envíos</p>
        </header>

        <div class="form-card">
            <form method="POST" enctype="multipart/form-data">
                <div class="form-group">
                    <label>📎 Archivo de Cuentas SMTP (Excel)</label>
                    <input type="file" name="archivo_cuentas" required accept=".xlsx,.xls">
                </div>

                <div class="form-group">
                    <label>👥 Archivo de Destinatarios (Excel)</label>
                    <input type="file" name="archivo_destinatarios" required accept=".xlsx,.xls">
                </div>

                <div class="form-group">
                    <label>📄 Archivo HTML del Cuerpo del Correo</label>
                    <input type="file" name="archivo_html" required accept=".html,.htm,.txt">
                </div>

                <div class="form-group">
                    <label>📋 Archivo de Asuntos (Excel o TXT)</label>
                    <input type="file" name="archivo_asuntos" required accept=".xlsx,.xls,.txt">
                </div>

                <div class="form-group">
                    <label>⚙️ Cantidad máxima de correos por cuenta</label>
                    <input type="number" name="max_correos" value="30" min="1" required>
                </div>

                <button type="submit">🚀 Iniciar Envío</button>
            </form>
        </div>

        {% if resultados %}
            <div class="results-card">
                <h2>✨ Resultados del Envío</h2>
                <ul class="results-list">
                {% for resultado in resultados %}
                    <li>{{ resultado }}</li>
                {% endfor %}
                </ul>
            </div>
        {% endif %}
    </div>
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

