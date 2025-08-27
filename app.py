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
TU_CORREO = ""
CONTRASENA_APLICACION = ""  # ej: "abcd efgh ijkl mnop"
ARCHIVO_EXCEL = 'correos.xlsx'
ASUNTO_CORREO = "DEMO WEEK (correo prueba automatizado)"

# --- PLANTILLA HTML FINAL CON IMÁGENES RESPONSIVE (DESKTOP/MOBILE) ---
def crear_plantilla_html(nombre_destinatario):
    # ¡IMPORTANTE! Reemplaza los '#' con tus enlaces reales.
    link_bildin = "#"
    link_talentum = "#"
    link_domus_ai = "#"
    link_quix = "#"
    link_vera = "#"
    link_nos = "#"
    link_virtual_demo_week = "#"

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>¡Ya formas parte de la Virtual Demo Week de UTEC Ventures!</title>
        <style>
            /* --- ESTILOS PARA INTERCAMBIO DE IMÁGENES --- */
            
            /* Ocultar el banner móvil por defecto */
            .mobile-banner {{
                display: none;
                max-height: 0;
                overflow: hidden;
            }}

            /* Media Query: Si la pantalla es de 600px o menos */
            @media screen and (max-width: 600px) {{
                /* Ocultar el banner de escritorio */
                .desktop-banner {{
                    display: none !important;
                }}
                /* Mostrar el banner de móvil */
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
                                <!-- Banner de Escritorio (Visible por defecto) -->
                                <div class="desktop-banner" style="margin: 0; padding: 0;">
                                    <a href="{link_virtual_demo_week}" target="_blank">
                                        <img src="cid:banner-desktop" alt="UTEC Ventures - Virtual Demo Week" width="600" style="display: block; width: 100%; max-width: 600px; height: auto; border-radius: 8px 8px 0 0;">
                                    </a>
                                </div>

                                <!-- Banner de Móvil (Oculto por defecto y oculto para Outlook) -->
                                <!--[if !mso]><!-->
                                <div class="mobile-banner" style="display:none;max-height:0;overflow:hidden;">
                                    <a href="{link_virtual_demo_week}" target="_blank">
                                        <img src="cid:banner-mobile" alt="UTEC Ventures - Virtual Demo Week" width="100%" style="display: block; width: 100%; max-width: 100%; height: auto; border-radius: 8px 8px 0 0;">
                                    </a>
                                </div>
                                <!--<![endif]-->
                            </td>
                        </tr>

                        <!-- SECCIÓN DEL CONTENIDO PRINCIPAL -->
                        <tr>
                            <td style="padding: 30px 30px 20px 30px; color: #333333; line-height: 1.6;">
                                <p style="margin: 0 0 15px 0;">Hola {nombre_destinatario},</p>
                                <p style="margin: 0 0 15px 0;">Del <strong>25 al 29 de agosto</strong> podrás entrar a nuestra plataforma para conocer a la nueva generación de startups en LatAm, explorar sus negocios y agendar reuniones uno a uno con los <strong>founders</strong>.</p>
                                <p style="margin: 0 0 25px 0;">Aquí un adelanto de lo que encontrarás:</p>
                                
                                <h2 style="color: #0089B1; margin: 0 0 10px 0;">Batch 14G – nuestras nuevas inversiones</h2>
                                <ul style="margin: 0 0 20px 0; padding-left: 20px; list-style-position: outside;">
                                    <li style="margin-bottom: 10px;">¿Quieres optimizar costos en obra con BIM + IA? Explora <a href="{link_bildin}" style="color: #330072; text-decoration: underline;"><strong>Bildin</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Buscas un recruiter de IA que trabaje 24/7? Descubre <a href="{link_talentum}" style="color: #330072; text-decoration: underline;"><strong>Talentum</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Te imaginas un edificio que se gestione solo? Conoce <a href="{link_domus_ai}" style="color: #330072; text-decoration: underline;"><strong>Domus AI</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Necesitas entrenamientos corporativos listos en 24h por WhatsApp? Revisa <a href="{link_quix}" style="color: #330072; text-decoration: underline;"><strong>Quix</strong></a>.</li>
                                </ul>
                                
                                <h2 style="color: #0089B1; margin: 0 0 10px 0;">Portafolio en crecimiento</h2>
                                <ul style="margin: 0 0 25px 0; padding-left: 20px; list-style-position: outside;">
                                    <li style="margin-bottom: 10px;">¿Quieres ver cómo se aceleran las decisiones de crédito 100x más rápido? Mira <a href="{link_vera}" style="color: #330072; text-decoration: underline;"><strong>VERA (Batch 13G)</strong></a>.</li>
                                    <li style="margin-bottom: 10px;">¿Te interesa cómo convertir finanzas informales en data crediticia trazable en zonas rurales? Aprende más de <a href="{link_nos}" style="color: #330072; text-decoration: underline;"><strong>NOS</strong></a>, startup del <strong>UV Lab</strong>.</li>
                                </ul>
                                
                                <p style="margin: 0 0 20px 0;">Ya puedes explorarlo todo desde un solo lugar. Accede aquí a la Virtual Demo Week</p>
                                
                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                        <td align="center">
                                            <a href="{link_virtual_demo_week}" style="background-color: #FF2A00; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">
                                                Acceder a la Plataforma
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        
                         <tr>
                           <td style="padding: 0px 30px 30px 30px; font-family: Arial, sans-serif; color: #333333; line-height: 1.6;">
                                <p style="margin: 20px 0 0 0;">El equipo de UTEC Ventures</p>
                           </td>
                        </tr>
                        
                        <tr>
                            <td align="center" style="padding: 20px; background-color: #f0f0f0; color: #555555; font-size: 12px; border-radius: 0 0 8px 8px;">
                                &copy; {pd.Timestamp.now().year} UTEC Ventures. Todos los derechos reservados.
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
            
            cuerpo_html_para_email = crear_plantilla_html(nombre)
            parte_html = MIMEText(cuerpo_html_para_email, "html")
            mensaje.attach(parte_html)
            
            # --- ADJUNTAR AMBOS BANNERS ---
            
            # 1. Adjuntar Banner de Escritorio
            ruta_banner_desktop = os.path.join('images', 'banner1.jpg')
            try:
                with open(ruta_banner_desktop, 'rb') as f:
                    parte_banner_desktop = MIMEImage(f.read())
                parte_banner_desktop.add_header('Content-ID', '<banner-desktop>')
                mensaje.attach(parte_banner_desktop)
            except FileNotFoundError:
                print(f"❌ ADVERTENCIA: No se encontró '{ruta_banner_desktop}'.")

            # 2. Adjuntar Banner de Móvil
            ruta_banner_mobile = os.path.join('images', 'banner2.jpg')
            try:
                with open(ruta_banner_mobile, 'rb') as f:
                    parte_banner_mobile = MIMEImage(f.read())
                parte_banner_mobile.add_header('Content-ID', '<banner-mobile>')
                mensaje.attach(parte_banner_mobile)
            except FileNotFoundError:
                print(f"❌ ADVERTENCIA: No se encontró '{ruta_banner_mobile}'.")
            
            # --- PREPARAR LA VISTA PREVIA ---
            ruta_desktop_web = ruta_banner_desktop.replace('\\', '/')
            ruta_mobile_web = ruta_banner_mobile.replace('\\', '/')
            cuerpo_html_para_preview = cuerpo_html_para_email.replace(
                'cid:banner-desktop', ruta_desktop_web
            ).replace(
                'cid:banner-mobile', ruta_mobile_web
            )
            
            filepath = "vista_previa_temporal.html"
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(cuerpo_html_para_preview)
            
            print("👀 Abriendo vista previa en tu navegador...")
            webbrowser.open("file://" + os.path.realpath(filepath))
            time.sleep(1)

            confirmacion = input(f"   ❓ ¿Enviar este correo a {nombre}? (s/n): ").lower()

            if confirmacion == 's':
                servidor.sendmail(TU_CORREO, correo_destino, mensaje.as_string())
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