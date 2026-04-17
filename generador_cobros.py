import os
import smtplib
import subprocess
from email.message import EmailMessage
from datetime import datetime
from dotenv import load_dotenv

from fpdf import FPDF
from database import SessionLocal
from models import Cliente

# Cargar variables de entorno
load_dotenv()

# Configuración SMTP (Outlook)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp-mail.outlook.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")

class CartaCobro(FPDF):
    def header(self):
        # Logo de la empresa
        logo_path = os.path.join("assets", "img", "logo.png")
        if os.path.exists(logo_path):
            self.image(logo_path, 18, 8, 48)
            
        self.set_font('Helvetica', '', 12)
        # JOSÉ A Y GERARDO E ZULUAGA S.A.S. (sin negrita en el PDF, o poco)
        self.cell(0, 5, 'José A y Gerardo E Zuluaga S.A.S.', new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_font('Helvetica', 'B', 9)
        self.cell(0, 5, 'IMPORTADORES DE GRANOS, ESPECIAS Y CONDIMENTOS', new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_font('Helvetica', 'B', 8)
        self.cell(0, 4, 'NIT: 890.928.717-5', new_x="LMARGIN", new_y="NEXT", align='C')
        self.ln(10)

    def footer(self):
        self.set_y(-25)
        self.set_font('Helvetica', '', 9)
        self.set_text_color(0, 102, 204) # Azul link
        self.cell(0, 5, "www.zuluagahermanos.com", new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_text_color(0, 0, 0)
        self.cell(0, 5, "CENTRAL MAYORISTA BLOQUE2-LC13 ITAGUI-ANTIOQUIA TEL: (604) 320 23 80", new_x="LMARGIN", new_y="NEXT", align='C')
        self.cell(0, 5, "CEL: 3172981578", new_x="LMARGIN", new_y="NEXT", align='C')

def enviar_correo(pdf_path, email_dest, nombre_cliente, total_mora, max_dias):
    if not SMTP_USER or not SMTP_PASS:
        print(f"⚠️ No hay credenciales SMTP configuradas en el .env (SMTP_USER, SMTP_PASS). Saltando envío de correo para {email_dest}")
        return False
        
    msg = EmailMessage()
    msg['Subject'] = "CARTA DE COBRO"
    msg['From'] = SMTP_USER
    msg['To'] = email_dest
    
    mora_str = f"{total_mora:,.0f}".replace(",", ".")
    cuerpo = f"""Estimada {nombre_cliente}

Nos dirigimos a usted con el fin de recordarle que a la fecha han transcurrido más {max_dias} días desde su vencimiento.

El monto adeudado es de ${mora_str}. 

Apreciamos mucho su atención a este asunto y agradeceríamos que nos informe si existe alguna dificultad con el pago o si requiere algún detalle adicional para proceder.

En caso de que el pago ya haya sido realizado, le pedimos que ignore este mensaje y acepte nuestras disculpas.

Quedamos atentos a su pronta respuesta y a la regularización de este saldo pendiente."""
    msg.set_content(cuerpo)
    
    with open(pdf_path, 'rb') as f:
        pdf_data = f.read()
        
    msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))
    
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"❌ Error al enviar correo a {email_dest}: {e}")
        return False

def generar(nit_especifico=None, db_session=None):
    db = db_session if db_session else SessionLocal()
    try:
        # Consultamos los clientes que tienen un saldo pendiente
        query = db.query(Cliente).filter(Cliente.total_cop > 0)
        if nit_especifico:
            query = query.filter(Cliente.nit_cliente == nit_especifico)
        clientes_db = query.all()
        
        # Agrupar facturas por NIT_CLIENTE
        agrupado = {}
        for c in clientes_db:
            nit = c.nit_cliente
            if nit not in agrupado:
                agrupado[nit] = []
            agrupado[nit].append(c)
            
        if not agrupado:
            print("No hay facturas con saldo pendiente (total_cop > 0) procedentes de la base de datos.")
            return

        out_dir = "cartas_generadas"
        os.makedirs(out_dir, exist_ok=True)
        
        for nit, facturas in agrupado.items():
            # Ordenar las facturas por fecha de vencimiento
            facturas = sorted(facturas, key=lambda x: x.fecha_vcto if x.fecha_vcto else datetime.min.date())
            primera_fac = facturas[0]
            nombre_cliente = primera_fac.razon_social or "CLIENTE DESCONOCIDO"
            correo = primera_fac.correo
            
            total_mora = sum([float(f.total_cop) for f in facturas if f.total_cop])
            
            print(f"\nLectura: Archivo de cartera detectado: {len(facturas)} facturas encontradas para {nombre_cliente}")
            
            pdf = CartaCobro()
            pdf.set_margins(18, 15, 18)
            pdf.add_page()
            
            # Fecha a la derecha
            meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            hoy = datetime.now()
            str_fecha = f"{hoy.day} de {meses[hoy.month-1]} de {hoy.year}"
            pdf.set_font('Helvetica', '', 10)
            pdf.cell(0, 10, str_fecha, new_x="LMARGIN", new_y="NEXT", align='R')
            pdf.ln(5)
            
            # Datos Cliente
            pdf.set_font('Helvetica', 'B', 10)
            pdf.cell(0, 4, f"{nombre_cliente}", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font('Helvetica', '', 9)
            
            # Evitar Nones
            tel = primera_fac.telefono or ""
            cel = primera_fac.celular or ""
            tels = " - ".join(filter(None, [tel, cel]))
            
            pdf.set_text_color(0, 51, 204) # Azul
            if tels:
                pdf.cell(0, 4, tels, new_x="LMARGIN", new_y="NEXT")
            if correo:
                pdf.cell(0, 4, f"{correo}", new_x="LMARGIN", new_y="NEXT")
            pdf.set_text_color(0, 0, 0) # Negro
            
            pdf.ln(5)
            # Asunto
            pdf.set_font('Helvetica', 'B', 10)
            pdf.cell(0, 8, "Asunto: Carta de Cobro.", new_x="LMARGIN", new_y="NEXT")
            
            pdf.ln(5)
            pdf.set_font('Helvetica', '', 10)
            w_est = pdf.get_string_width("Estimado ")
            pdf.cell(w_est, 5, "Estimado ", align='L')
            pdf.set_font('Helvetica', 'B', 10)
            pdf.cell(0, 5, f"{nombre_cliente}", new_x="LMARGIN", new_y="NEXT", align='L')
            
            pdf.ln(5)
            pdf.set_font('Helvetica', '', 10)
            
            mora_str = f"{total_mora:,.0f}".replace(",", ".")
            intro_text = f"A la fecha de hoy su cuenta presenta un saldo vencido de $ {mora_str}, correspondiente a las facturas"
            pdf.multi_cell(0, 5, intro_text)
            pdf.ln(5)
            
            # Tabla
            pdf.set_font('Helvetica', '', 9)
            pdf.set_fill_color(255, 255, 255)
            
            # Anchos: Total 174 mm para cuadrar con márgenes de 18 (210 - 36 = 174)
            w1, w2, w3, w4, w5 = 62, 28, 25, 25, 34
            pdf.cell(w1, 6, 'Razón social', border=1)
            pdf.cell(w2, 6, 'Nro. docto. cruce', border=1)
            pdf.cell(w3, 6, 'Días vencidos', border=1)
            pdf.cell(w4, 6, 'Fecha vcto.', border=1)
            pdf.cell(w5, 6, 'Total COP', border=1, new_x="LMARGIN", new_y="NEXT")
            
            # Filas
            for f in facturas:
                f_doc = str(f.nro_docto_cruce)
                f_vcto = f.fecha_vcto.strftime("%d/%m/%Y") if f.fecha_vcto else "N/A"
                f_dias = str(f.dias_vencidos) if f.dias_vencidos is not None else "0"
                saldo_val = float(f.total_cop) if f.total_cop else 0.0
                f_saldo = f"{saldo_val:,.0f}".replace(",", ".")
                
                pdf.cell(w1, 6, str(nombre_cliente)[:35], border=1)
                pdf.cell(w2, 6, f_doc, border=1)
                pdf.cell(w3, 6, f_dias, border=1)
                pdf.cell(w4, 6, f_vcto, border=1)
                pdf.cell(w5, 6, f_saldo, border=1, new_x="LMARGIN", new_y="NEXT")
                
            pdf.ln(8)
            
            # Texto Legal
            pdf.set_font('Helvetica', '', 9.5)
            texto_legal1 = "Tenga presente que la ley 1266 de Habeas Data y las nuevas disposiciones en materia crediticia, exigen a JOSÉ A Y GERARDO E ZULUAGA S.A.S. el reporte a las centrales de riesgo (Datacredito), al igual que a sus codeudores."
            
            texto_html2 = """Para evitar ser reportado negativamente, comuníquese al teléfono <font color="#0033cc">(604) 320 23 80</font> celular <font color="#0033cc">312 833 4630</font> los correos electrónicos <font color="#0033cc"><a href="mailto:credito-cartera@josegera.com">credito-cartera@josegera.com</a> ; <a href="mailto:cartera@josegera.com">cartera@josegera.com</a></font> para establecer acuerdo de pago."""
            
            texto_legal3 = "Si al recibir esta carta usted se encuentra al día, por favor haga caso omiso a su contenido."
            
            pdf.multi_cell(0, 5, texto_legal1)
            pdf.ln(4)
            pdf.write_html(texto_html2)
            pdf.ln(4)
            pdf.multi_cell(0, 5, texto_legal3)
            
            pdf.ln(8)
            
            # Firma (la imagen ya incluye Atentamente y datos)
            firma_path = os.path.join("assets", "img", "firma_sebastian.jpeg")
            if os.path.exists(firma_path):
                y_firma = pdf.get_y()
                pdf.image(firma_path, 18, y_firma, 60)
                pdf.ln(35)
            else:
                pdf.ln(15)
            
            # Guardar temporalmente con nombre limpio
            nombre_safe = "".join([c for c in nombre_cliente if c.isalnum() or c==' ']).strip()
            base_pdf_path = os.path.join(out_dir, f"Cobro_{nombre_safe}.pdf")
            pdf_path = base_pdf_path
            counter = 1
            while True:
                try:
                    pdf.output(pdf_path)
                    break
                except PermissionError:
                    pdf_path = os.path.join(out_dir, f"Cobro_{nombre_safe} ({counter}).pdf")
                    counter += 1
            
            # Envío o apertura manual
            if correo:
                print(f"Validación: Email encontrado: {correo}. Procediendo a envío.")
                max_dias_cliente = max([int(f.dias_vencidos) for f in facturas if f.dias_vencidos is not None], default=0)
                enviado = enviar_correo(pdf_path, correo, nombre_cliente, total_mora, max_dias_cliente)
                if enviado:
                    print("Resultado: PDF generado y enviado exitosamente. Registro procesado en BD.")
                else:
                    print("Resultado: PDF generado pero hubo un error en la conexión del correo.")
            else:
                print(f"Validación: No se encontró email para {nombre_cliente}. Se abrirá el archivo generado.")
                abs_path = os.path.abspath(pdf_path)
                try:
                    # Comando nativo de Windows (explorer) para seleccionar el archivo en carpeta
                    subprocess.Popen(f'explorer /select,"{abs_path}"')
                except Exception as e:
                    print(f"No se pudo abrir el explorador de forma automática: {e}")
                    
    finally:
        if not db_session:
            db.close()

if __name__ == "__main__":
    generar()
