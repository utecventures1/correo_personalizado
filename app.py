import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import webbrowser
import os
import time
from email.mime.image import MIMEImage

# --- CONFIGURACIÓN ---
# ¡Completa estos datos!
TU_CORREO = "dbazan@utec.edu.pe"
CONTRASENA_APLICACION = ""  # ej: "abcd efgh ijkl mnop"
ARCHIVO_EXCEL = 'correos.xlsx'
ASUNTO_CORREO = "Conoce DomusAI: Una plataforma para la gestión de edificios automatizada con IA"
COPIA_A = [] 
COPIA_OCULTA_A = ["pkingkee@utec.edu.pe"]

# --- PLANTILLA HTML ACTUALIZADA ---
# --- PLANTILLA HTML FINAL CON TODOS LOS CAMBIOS ---
def crear_plantilla_html(nombre_destinatario):
    # Enlaces actualizados
    link_pitch_domus = "https://www.youtube.com/watch?v=wMrrHkjcowk"
    link_reunion_founders = "https://calendly.com/domus-ai/demo?back=1&month=2025-08"
    # El link del banner y del texto apuntan al mismo sitio para consistencia
    link_demo_day = "https://eventos.utec.edu.pe/DemoDay"

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale-1.0">
        <title>Conoce DomusAI</title>
        <style>
            /* --- ESTILOS PARA INTERCAMBIO DE IMÁGENES --- */
            .mobile-banner {{
                display: none;
                max-height: 0;
                overflow: hidden;
            }}
            @media screen and (max-width: 600px) {{
                .desktop-banner {{ display: none !important; }}
                .mobile-banner {{
                    display: block !important;
                    max-height: none !important;
                    overflow: visible !important;
                    width: 100% !important;
                }}
            }}
        </style>
    </head>
    <body style="margin: 0; padding: 0; background-color: #f4f4f4; font-family: Arial, sans-serif;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
                <td style="padding: 20px 10px;">
                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border-collapse: collapse; background-color: #ffffff; border-radius: 8px; max-width: 600px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                        
                        <!-- SECCIÓN DEL BANNER RESPONSIVE -->
                        <tr>
                            <td align="center">
                                <div class="desktop-banner" style="margin: 0; padding: 0;">
                                    <a href="{link_demo_day}" target="_blank">
                                        <img src="cid:banner-desktop" alt="UTEC Ventures Demo Day" width="600" style="display: block; width: 100%; max-width: 600px; height: auto; border-radius: 8px 8px 0 0;">
                                    </a>
                                </div>
                                <!--[if !mso]><!-->
                                <div class="mobile-banner" style="display:none;max-height:0;overflow:hidden;">
                                    <a href="{link_demo_day}" target="_blank">
                                        <img src="cid:banner-mobile" alt="UTEC Ventures Demo Day" width="100%" style="display: block; width: 100%; max-width: 100%; height: auto; border-radius: 8px 8px 0 0;">
                                    </a>
                                </div>
                                <!--<![endif]-->
                            </td>
                        </tr>

                        <!-- SECCIÓN DEL CONTENIDO PRINCIPAL -->
                        <tr>
                            <td style="padding: 30px 30px 20px 30px; color: #333333; line-height: 1.6;">
                                <p style="margin: 0 0 15px 0;">Hola <strong>{nombre_destinatario}</strong>,</p>
                                <p style="margin: 0 0 15px 0;">¿Te imaginas un edificio que se gestione solo?</p>
                                <p style="margin: 0 0 25px 0;">En UTEC Ventures seguimos apostando por startups con alto potencial. Hoy queremos presentarte a <strong>DomusAI</strong>, parte de nuestra 14G.</p>
                                <p style="margin: 0 0 25px 0;">
                                    <strong>DomusAI</strong> automatiza la facturación, el mantenimiento y la atención a residentes para los administradores de propiedades, mediante agentes de IA en WhatsApp y llamadas telefónicas. Su plataforma ayuda a los equipos a reducir carga operativa, mejorar el flujo de caja y brindar servicio inmediato (sin necesidad de contratar más personal ni capacitarlo).
                                </p>
                                
                                <ul style="margin: 0 0 20px 0; padding-left: 20px; list-style-position: outside;">
                                    <li style="margin-bottom: 10px;">¿Quieres ver cómo <strong>DomusAI</strong> está transformando la gestión inmobiliaria? Mira aquí su <a href="{link_pitch_domus}" target="_blank" style="color: #330072; text-decoration: underline;"><strong>pitch</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Te interesa conversar directamente con los founders? Agenda una reunión <a href="{link_reunion_founders}" target="_blank" style="color: #330072; text-decoration: underline;"><strong>aquí</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Quieres conocer más startups de nuestra 14G? <a href="{link_demo_day}" target="_blank" style="color: #330072; text-decoration: underline;"><strong>Explora la 14G completa</strong></a>.</li>
                                </ul>

                                <p style="margin: 0 0 20px 0;">Estamos convencidos del potencial de <strong>DomusAI</strong> para convertirse en el sistema operativo de edificios impulsado por IA y te invitamos a conocerla más de cerca.</p>
                            </td>
                        </tr>
                        
                         <tr>
                           <td style="padding: 0px 30px 30px 30px; font-family: Arial, sans-serif; color: #333333; line-height: 1.6;">
                                <p style="margin: 0 0 0 0;">Saludos,</p>
                                <p style="margin: 0;">Equipo de UTEC Ventures</p>
                           </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
    return html

def main():
    try:
        df = pd.read_excel(ARCHIVO_EXCEL)
        
        contexto_ssl = ssl.create_default_context()
        servidor = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto_ssl)
        servidor.login(TU_CORREO, CONTRASENA_APLICACION)
        print("✅ Conexión exitosa con Gmail.")

        for index, row in df.iterrows():
            nombre = row['nombre']
            correo_destino = row['correo']
            
            print(f"\n--------------------------------------------------")
            print(f"Preparando correo para: {nombre} <{correo_destino}>")

            mensaje = MIMEMultipart('related')
            mensaje["Subject"] = ASUNTO_CORREO
            mensaje["From"] = TU_CORREO
            mensaje["To"] = correo_destino
            
            if COPIA_A:
                mensaje["Cc"] = ", ".join(COPIA_A)

            cuerpo_html_para_email = crear_plantilla_html(nombre)
            parte_html = MIMEText(cuerpo_html_para_email, "html")
            mensaje.attach(parte_html)
            
            # --- ADJUNTAR AMBOS BANNERS ---
            try:
                with open(os.path.join('images', 'banner1.jpg'), 'rb') as f:
                    parte_banner_desktop = MIMEImage(f.read())
                parte_banner_desktop.add_header('Content-ID', '<banner-desktop>')
                mensaje.attach(parte_banner_desktop)
            except FileNotFoundError:
                print("❌ ADVERTENCIA: No se encontró 'images/banner1.jpg'. El banner de escritorio no se adjuntará.")

            try:
                with open(os.path.join('images', 'banner2.jpg'), 'rb') as f:
                    parte_banner_mobile = MIMEImage(f.read())
                parte_banner_mobile.add_header('Content-ID', '<banner-mobile>')
                mensaje.attach(parte_banner_mobile)
            except FileNotFoundError:
                print("❌ ADVERTENCIA: No se encontró 'images/banner2.jpg'. El banner móvil no se adjuntará.")
            
            # --- PREPARAR LA VISTA PREVIA ---
            cuerpo_html_para_preview = cuerpo_html_para_email.replace(
                'cid:banner-desktop', os.path.join('images', 'banner1.jpg').replace('\\', '/')
            ).replace(
                'cid:banner-mobile', os.path.join('images', 'banner2.jpg').replace('\\', '/')
            )
            
            filepath = "vista_previa_temporal.html"
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(cuerpo_html_para_preview)
            
            print("👀 Abriendo vista previa en tu navegador...")
            webbrowser.open("file://" + os.path.realpath(filepath))
            time.sleep(1)
            
            print(f"   - Con Copia (CC) a: {', '.join(COPIA_A) if COPIA_A else 'Ninguno'}")
            print(f"   - Con Copia Oculta (CCO) a: {', '.join(COPIA_OCULTA_A) if COPIA_OCULTA_A else 'Ninguno'}")
            
            confirmacion = input(f"   ❓ ¿Enviar este correo a {nombre}, su correo es {correo_destino}? (s/n): ").lower()

            if confirmacion == 's':
                destinatarios_completos = [correo_destino] + COPIA_A + COPIA_OCULTA_A
                servidor.sendmail(TU_CORREO, destinatarios_completos, mensaje.as_string())
                print(f"   🚀 ¡Correo enviado exitosamente!")
            else:
                print(f"   ❌ Envío cancelado por el usuario.")
            
            os.remove(filepath)

        servidor.quit()
        print("\n🎉 Proceso completado.")

    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")

if __name__ == '__main__':
    main()
