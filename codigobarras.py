##             (               (              (           (                            
##             )\      )  (    )\             )\        ( )\     (   (   (             
##           (((_)  ( /(  )(  ((_) (   (   ((((_)(      )((_)   ))\  )\  )(    (   (   
##           )\___  )(_))(()\  _   )\  )\   )\ _ )\    ((_)_   /((_)((_)(()\   )\  )\  
##          ((/ __|((_)_  ((_)| | ((_)((_)  (_)_\(_)    / _ \ (_))(  (_) ((_) ((_)((_) 
##           | (__ / _` || '_|| |/ _ \(_-<   / _ \  _  | (_) || || | | || '_|/ _ \|_ / 
##            \___|\__,_||_|  |_|\___//__/  /_/ \_\(_)  \__\_\ \_,_| |_||_|  \___//__| 
                                                                           
import barcode
from barcode.writer import ImageWriter
from datetime import datetime
from PIL import Image
import io
import win32clipboard  # Necesario: pip install pywin32

# Solicitar el número consecutivo al usuario
consecutivo = input("Ingrese el número consecutivo: ").strip()  # Se mantiene como cadena

# Datos para el código de barras
uadmva = "22803002L/"
anoactual = f"/{datetime.now().year}"
valor_codbarra = f"{uadmva}{consecutivo}{anoactual}"

# Configurar opciones de tamaño personalizadas
options = {
    "module_width": 0.3,   # Ancho de las barras
    "module_height": 10,   # Altura del código de barras
    "font_size": 10,       # Tamaño del texto debajo
    "text_distance": 5,    # Distancia entre barras y texto
}

# Generar código de barras con tamaño personalizado
code128 = barcode.get_barcode_class('code128')
barcode_instance = code128(valor_codbarra, writer=ImageWriter())

# Guardar la imagen con opciones personalizadas
filename = f"cod_barras_{consecutivo}"
full_path = f"{filename}.png"
barcode_instance.save(filename, options=options)

# Abrir la imagen recién creada
image = Image.open(full_path)

# Convertir la imagen a formato BMP para copiar al portapapeles
output = io.BytesIO()
image.convert("RGB").save(output, "BMP")
data = output.getvalue()[14:]  # Saltar cabecera BMP
output.close()

# Copiar al portapapeles
win32clipboard.OpenClipboard()
win32clipboard.EmptyClipboard()
win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
win32clipboard.CloseClipboard()

print(f"Código de barras generado: {full_path} y copiado al portapapeles.")