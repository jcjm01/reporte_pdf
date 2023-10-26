
import pymysql
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# En este apartado se configura la conexión de la BD con Python
conector = pymysql.connect(
    host="aqui_va_tu_host",
    user="aqui_va_tu_usuario",
    password="aqui_va_tu_contraseña_de_BD",
    database="aqui_va_tu_BD",
    charset='utf8'
)
#########
try:
    # Aquí creamos un cursor para hacer las consultas a la BD
    cursor = conector.cursor()

    # Este query es un INNER JOIN que consulta tres tablas en MySQL para generar el reporte
    query = "SELECT cliente.IdCliente,cliente.Departamento,maquina.Capacidad,maquina.RAM,maquina.ESTADO,vcenters.IP,vcenters.ESTATUS FROM cliente INNER JOIN maquina ON cliente.IdCliente = maquina.IdCliente  INNER JOIN vcenters ON cliente.IdCliente= vcenters.IdCliente WHERE 'IdCliente'<= 10;"

    # Ejecuta la consulta en MySQL
    cursor.execute(query)

    # Guarda los resultados de la query y los guarda en un DataFrame de la librería Pandas
    resultados = cursor.fetchall()
    column_names = [i[0] for i in cursor.description]
    df = pd.DataFrame(resultados, columns=column_names)

    # Guarda el DataFrame anterior en un archivo PDF
    archivo_pdf = "resultados.pdf"
    """
    1.-for column in df.columns:: Este bucle itera a través de las columnas del DataFrame df. df.columns es una lista de los nombres de las columnas en el DataFrame.
    2.-pdf.cell(30, 8, column, border=1): Aquí se utiliza el método cell de la instancia de la clase pdf para agregar una celda a la tabla en el PDF. Los valores tienen el siguiente significado:

    40: Es el ancho de la celda en unidades de medida (generalmente en milímetros en el caso predeterminado de fpdf).
    10: Es la altura de la celda en unidades de medida.
    column: Es el contenido que se inserta en la celda. En este caso, es el nombre de una columna del DataFrame.
    border=1: Indica que se deben dibujar bordes alrededor de la celda.

    3.-pdf.ln(): Después de agregar una fila de celdas correspondiente a las columnas, se utiliza ln para mover el cursor a la siguiente línea (fila) de la tabla.
    
    4.-for index, row in df.iterrows():: Este bucle itera a través de las filas del DataFrame df. df.iterrows() devuelve un generador que produce índices de fila y las filas correspondientes.
    
    5.-for col in row:: Este bucle itera a través de las celdas (valores) en una fila del DataFrame.

    6.-pdf.cell(30, 8, str(col), border=1): Al igual que en el bucle anterior, se utiliza cell para agregar una celda a la tabla en el PDF. Los valores tienen el siguiente significado:

    40: Ancho de la celda.
    10: Altura de la celda.
    str(col): El contenido que se inserta en la celda. Se convierte a una cadena (string) para asegurarse de que el contenido sea legible en el PDF.
    border=1: Se especifica que se deben dibujar bordes alrededor de la celda.

    7.-pdf.ln(): Nuevamente, después de agregar una fila de celdas correspondiente a los valores de una fila, se utiliza ln para mover el cursor a la siguiente línea (fila) de la tabla.
    """
    class PDF(FPDF):
        def header(self):
            self.set_font("Arial", "B", 12)
            self.cell(0, 10, "Reporte de la BD", align="C", ln=True)

        def footer(self):
            self.set_y(-15)
            self.set_font("Arial", "I", 8)
            self.cell(0, 10, f"Página {self.page_no()}", align="C")

    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Arial", size=9)

    for column in df.columns:
        pdf.cell(30, 8, column, border=1)
    pdf.ln()

    for index, row in df.iterrows():
        for col in row:
            pdf.cell(30,8, str(col), border=1)
        pdf.ln()

    pdf.output(archivo_pdf)
    #
    # En este apartado se configura el envío de correo
    envia_email = "aqui_va_el_correo_de_quien_envia_el_mail"
    recibe_email = "aqui_va_el_mail_de_quien_recibe_el_mail"
    

    asunto = "Correo de prueba con archivo PDF adjunto"
    cuerpo_email = "Email de prueba de envío automatizado de correos que se ejecuta una vez al día adjuntando un PDF con un reporte generado de una BD en MySQL."

    # Arma el mensaje del correo con los destinatarios y el que envía el correo
    msg = MIMEMultipart()
    msg["From"] = envia_email
    msg["To"] = recibe_email
    msg["Subject"] = asunto

    # Anexa el cuerpo del correo
    msg.attach(MIMEText(cuerpo_email, "plain"))

    # Este bloque adjunta el archivo PDF generado en la consulta INNER JOIN
    with open(archivo_pdf, "rb") as anexar_pdf:
        part = MIMEApplication(anexar_pdf.read())
        part.add_header("Content-Disposition", f"attachment; filename={archivo_pdf}")
        msg.attach(part)

    # En este bloque se configura el servidor SMTP con el puerto del mismo y el correo desde donde lo enviamos
    smtp_server = "smtp-mail.outlook.com"
    smtp_puerto = 587
    smtp_usuario = "aqui_va_tu_correo"
    smtp_passwd = "aqui_va_tu_contraseña_de_tu_correo"

    # Este bloque inicia el servicio SMTP
    servidorsmtp = smtplib.SMTP(smtp_server, smtp_puerto)
    servidorsmtp.starttls()

    # Este bloque inicia sesión del servidor SMTP
    servidorsmtp.login(smtp_usuario, smtp_passwd)

    # Aquí se envía el correo electrónico
    servidorsmtp.sendmail(envia_email, recibe_email, msg.as_string())

    # Este bloque cierra la conexión con el servidor SMTP
    servidorsmtp.quit()

# Se usó el método try-except para que, en caso de fallar, el programa no deje de correr pero muestre un mensaje de error
except Exception as e:
    print(f"Error: {e}")

finally:
    # Este módulo termina el cursor y la conexión
    cursor.close()
    conector.close()
########