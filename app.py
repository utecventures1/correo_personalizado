import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import webbrowser
import os
import time

# --- CONFIGURACIÓN ---
# ¡Completa estos datos!
TU_CORREO = ""
CONTRASENA_APLICACION = "" # ej: "abcd efgh ijkl mnop"
ARCHIVO_EXCEL = 'correos.xlsx'
ASUNTO_CORREO = "Un Asunto Verificado para Ti"

def crear_plantilla_html(nombre_destinatario):
    """Crea el cuerpo del correo en formato HTML."""
    # (Pega aquí la misma función crear_plantilla_html que te di antes)
    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Comunicado UTEC Ventures</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f4f4f4; }}
            .container {{ max-width: 600px; margin: auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}
            .header {{ background-color: #008C95; color: #ffffff; padding: 20px; text-align: center; border-radius: 8px 8px 0 0;}}
            .content {{ padding: 30px; line-height: 1.6; }}
            .button {{ display: inline-block; background-color: #FF481A; color: #ffffff; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; }}
            .footer {{ background-color: #f0f0f0; color: #555555; text-align: center; padding: 20px; font-size: 12px; border-radius: 0 0 8px 8px;}}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header"><h1>UTEC Ventures</h1></div>
            <div class="content">
                <p>Hola {nombre_destinatario},</p>
                <p>Este es el contenido del correo. Revisa que todo esté correcto antes de confirmar el envío.</p>
                <a href="https://www.utec.edu.pe/utec-ventures" class="button">Visita Nuestra Web</a>
            </div>
            <div class="footer"><p>&copy; {pd.Timestamp.now().year} UTEC Ventures</p></div>
        </div>
    </body>
    </html>
    """
    return html

# --- FUNCIÓN PRINCIPAL ---
def main():
    try:
        df = pd.read_excel(ARCHIVO_EXCEL)
        
        # Conexión segura con el servidor de Gmail
        contexto_ssl = ssl.create_default_context()
        servidor = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto_ssl)
        servidor.login(TU_CORREO, CONTRASENA_APLICACION)
        print("✅ Conexión exitosa con Gmail.")

        for index, row in df.iterrows():
            nombre = row['nombre']
            correo_destino = row['correo']
            
            print(f"\n--------------------------------------------------")
            print(f"Preparando correo para: {nombre} <{correo_destino}>")

            # 1. Crear el HTML personalizado
            cuerpo_html = crear_plantilla_html(nombre)
            
            # 2. Guardar el HTML en un archivo temporal para la vista previa
            filepath = "vista_previa_temporal.html"
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(cuerpo_html)

            # 3. Abrir la vista previa en el navegador
            print("👀 Abriendo vista previa en tu navegador...")
            webbrowser.open("file://" + os.path.realpath(filepath))
            time.sleep(1) # Pequeña pausa para que el navegador se abra

            # 4. Pedir confirmación al usuario
            confirmacion = input(f"   ❓ ¿Enviar este correo a {nombre}? (s/n): ").lower()

            # 5. Enviar (o no) según la respuesta
            if confirmacion == 's':
                mensaje = MIMEMultipart("alternative")
                mensaje["Subject"] = ASUNTO_CORREO
                mensaje["From"] = TU_CORREO
                mensaje["To"] = correo_destino
                mensaje.attach(MIMEText(cuerpo_html, "html"))
                
                servidor.sendmail(TU_CORREO, correo_destino, mensaje.as_string())
                print(f"   🚀 ¡Correo enviado exitosamente!")
            else:
                print(f"   ❌ Envío cancelado por el usuario. Saltando a la siguiente persona.")
            
            # Limpiar el archivo temporal
            os.remove(filepath)

        servidor.quit()
        print("\n--------------------------------------------------")
        print("🎉 Proceso completado. Todos los contactos han sido procesados.")

    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{ARCHIVO_EXCEL}'.")
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación. Revisa tu correo y la Contraseña de Aplicación.")
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado: {e}")

if __name__ == '__main__':
    main()