import os
import subprocess
from datetime import datetime
from fpdf import FPDF
from database import SessionLocal
from models import Cliente

class CertificadoComercial(FPDF):
    def header(self):
        # Logo de la empresa
        logo_path = os.path.join("assets", "img", "logo.png")
        if os.path.exists(logo_path):
            self.image(logo_path, 15, 8, 33)
            
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 5, 'José A y Gerardo E Zuluaga S.A.S.', new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_font('helvetica', 'B', 10)
        self.cell(0, 5, 'IMPORTADORES DE GRANOS, ESPECIAS Y CONDIMENTOS', new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_font('helvetica', 'B', 9)
        self.cell(0, 5, 'NIT: 890.928.717-5', new_x="LMARGIN", new_y="NEXT", align='C')
        self.ln(20)

    def footer(self):
        self.set_y(-25)
        self.set_font('helvetica', '', 9)
        self.set_text_color(0, 0, 255)
        self.cell(0, 5, 'www.zuluagahermanos.com', new_x="LMARGIN", new_y="NEXT", align='C', link='http://www.zuluagahermanos.com')
        self.set_text_color(0, 0, 0)
        self.cell(0, 5, 'CENTRAL MAYORISTA BLOQUE2-LC13 ITAGUI-ANTIOQUIA TEL: (604) 320 23 80', new_x="LMARGIN", new_y="NEXT", align='C')
        self.cell(0, 5, 'CEL: 3172981578', new_x="LMARGIN", new_y="NEXT", align='C')

def generar(nit_especifico, db_session=None):
    db = db_session if db_session else SessionLocal()
    try:
        cliente = db.query(Cliente).filter(Cliente.nit_cliente == nit_especifico).first()
        if not cliente:
            print(f"Cliente con NIT {nit_especifico} no encontrado.")
            return

        out_dir = "certificados_generados"
        os.makedirs(out_dir, exist_ok=True)
        
        pdf = CertificadoComercial()
        pdf.set_margins(15, 15, 15)
        pdf.add_page()
        
        # Fecha a la izquierda
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        hoy = datetime.now()
        str_fecha = f"Itagüí (Antioquia), {hoy.day} de {meses[hoy.month-1]} de {hoy.year}"
        
        pdf.set_font('helvetica', '', 11)
        pdf.cell(0, 6, str_fecha, new_x="LMARGIN", new_y="NEXT", align='L')
        pdf.ln(15)
        
        # A QUIEN PUEDA INTERESAR
        pdf.set_font('helvetica', 'B', 12)
        pdf.cell(0, 6, "A QUIEN PUEDA INTERESAR", new_x="LMARGIN", new_y="NEXT", align='C')
        pdf.ln(15)
        
        # Texto principal
        razon_social = cliente.razon_social or "CLIENTE DESCONOCIDO"
        pdf.set_font('helvetica', '', 11)
        texto_intro = "JOSE A. Y GERARDO E. ZULUAGA S.A.S. certifica que:"
        pdf.multi_cell(0, 6, texto_intro, align='J')
        pdf.ln(10)
        
        # Helper for currency
        def fmt_money(val):
            if val is None: return "$ 0"
            return f"$ {float(val):,.0f}".replace(",", ".")
            
        def fmt_date(d):
            if not d: return "N/A"
            return d.strftime("%d/%m/%Y")
            
        # Tabla de datos simulando columnas
        pdf.set_font('helvetica', 'B', 11)
        
        datos = [
            ("Cliente:", razon_social),
            ("Nit / Cliente:", cliente.nit_cliente or "N/A"),
            ("Cliente desde:", fmt_date(cliente.fecha_ingreso)),
            ("Condición de pago:", cliente.condicion_pago or "N/A"),
            ("Cupo de crédito:", fmt_money(cliente.cupo_credito)),
            ("Fecha última compra:", fmt_date(cliente.fecha_ultima_compra)),
            ("Promedio mensual:", fmt_money(cliente.promedio_mensual))
        ]
        
        for label, value in datos:
            pdf.set_font('helvetica', 'B', 11)
            pdf.cell(50, 8, label, new_x="RIGHT")
            pdf.set_font('helvetica', '', 11)
            pdf.cell(0, 8, str(value), new_x="LMARGIN", new_y="NEXT")
            
        pdf.ln(15)
        
        # Texto legal
        texto_legal = "Los anteriores datos son exclusivamente de carácter informativo y no comprometen a JOSE A. Y GERARDO E. ZULUAGA S.A.S. en la toma de decisiones de terceros ni en el uso que se les dé posteriormente."
        pdf.multi_cell(0, 6, texto_legal, align='J')
        pdf.ln(20)
        
        # Firma
        pdf.cell(0, 6, "Cordialmente,", new_x="LMARGIN", new_y="NEXT")
        firma_y = pdf.get_y()
        
        firma_path_1 = os.path.join("assets", "img", "image.png")
        firma_path_2 = os.path.join("assets", "img", "firma_sebastian (1).jpeg")
        
        if os.path.exists(firma_path_1):
            pdf.image(firma_path_1, 15, firma_y, w=40)
        elif os.path.exists(firma_path_2):
            pdf.image(firma_path_2, 15, firma_y, w=40)
            
        pdf.ln(20) 
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 5, "Contacto: Sebastián Álvarez Sepúlveda", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, "Email: credito_cartera@josegera.com", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, "Cel/WhatsApp: 310 884 1486", new_x="LMARGIN", new_y="NEXT")
        
        # Guardado
        nombre_safe = "".join([c for c in razon_social if c.isalnum() or c==' ']).strip()
        timestamp = hoy.strftime("%Y%m%d_%H%M%S")
        base_pdf_path = os.path.join(out_dir, f"Certificado_{nombre_safe}_{timestamp}.pdf")
        
        pdf.output(base_pdf_path)
        
        abs_path = os.path.abspath(base_pdf_path)
        try:
            subprocess.Popen(f'explorer /select,"{abs_path}"')
        except Exception as e:
            print(f"No se pudo abrir el explorador de forma automática: {e}")
            
        return base_pdf_path
        
    finally:
        if not db_session:
            db.close()

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        generar(sys.argv[1])
    else:
        print("Debe proveer un NIT específico.")
