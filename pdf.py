import os
import sys
import time
from PIL import Image
from PIL import JpegImagePlugin
from pypdf import PdfWriter, PdfReader
from pdf2docx import Converter
from docx2pdf import convert

os.system('')
CYAN = '\033[96m'
ROJO = '\033[91m'
VERDE = '\033[92m'
AMARILLO = '\033[93m'
BLANCO = '\033[97m'
RESET = '\033[0m'

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_ENTRADA = os.path.join(DIRECTORIO_BASE, "Entrada_PDF")
CARPETA_SALIDA = os.path.join(DIRECTORIO_BASE, "Salida_PDF")

def preparar_carpetas():
    for carpeta in [CARPETA_ENTRADA, CARPETA_SALIDA]:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)

def mostrar_banner():
    print(CYAN + r"""
  ███████╗██████╗  ██████╗ ███████╗    ██████╗  ██████╗  ██████╗ ███████╗
  ██╔════╝██╔══██╗██╔═══██╗██╔════╝    ██╔══██╗██╔═══██╗██╔════╝ ██╔════╝
  █████╗  ██████╔╝██║   ██║███████╗    ██║  ██║██║   ██║██║      ███████╗
  ██╔══╝  ██╔══██╗██║   ██║╚════██║    ██║  ██║██║   ██║██║      ╚════██║
  ███████╗██║  ██║╚██████╔╝███████║    ██████╔╝╚██████╔╝╚██████╗ ███████║
  ╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝    ╚═════╝  ╚═════╝  ╚═════╝ ╚══════╝""" + RESET)
    print(BLANCO + "    =======================================================================")
    print("                      By: Gianmarco Joel Miñano Peña")
    print("    =======================================================================\n" + RESET)

def listar_archivos(extension='.pdf'):
    archivos = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(extension) and not f.startswith('~$')]
    if not archivos:
        print(ROJO + f"[-] Error: No hay archivos {extension} en la carpeta 'Entrada_PDF'." + RESET)
        return None
    
    print(BLANCO + f"\nArchivos {extension} disponibles:" + RESET)
    for i, archivo in enumerate(archivos, 1):
        print(f"  {i}. {archivo}")
        
    while True:
        try:
            seleccion = int(input(CYAN + "\nIngrese el número del archivo que desea procesar: " + RESET))
            if 1 <= seleccion <= len(archivos):
                return os.path.join(CARPETA_ENTRADA, archivos[seleccion - 1])
            else:
                print(ROJO + "[-] Número fuera de rango. Intente nuevamente." + RESET)
        except ValueError:
            print(ROJO + "[-] Por favor, ingrese un número válido." + RESET)

def procesar_rangos(texto_rangos, total_paginas):
    grupos = []
    partes = texto_rangos.split(',')
    
    for parte in partes:
        parte = parte.strip()
        if '-' in parte:
            inicio, fin = parte.split('-')
            inicio, fin = int(inicio) - 1, int(fin) - 1
            grupos.append(list(range(inicio, fin + 1)))
        else:
            grupos.append([int(parte) - 1])
            
    return grupos

def unir_pdfs():
    print(AMARILLO + "\n[+] UNIR PDFs" + RESET)
    archivos = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith('.pdf')]
    
    if len(archivos) < 2:
        print(ROJO + "[-] Error: Necesitas al menos 2 PDFs en la carpeta 'Entrada_PDF' para unirlos." + RESET)
        return

    print(f"[*] Se detectaron {len(archivos)} archivos. Uniendo en orden alfabético...")
    fusionador = PdfWriter()
    
    for archivo in sorted(archivos):
        ruta_completa = os.path.join(CARPETA_ENTRADA, archivo)
        fusionador.append(ruta_completa)
        print(f"    -> Añadido: {archivo}")

    ruta_salida = os.path.join(CARPETA_SALIDA, f"EROS_Fusionado_{int(time.time())}.pdf")
    with open(ruta_salida, "wb") as salida:
        fusionador.write(salida)
    
    print(VERDE + f"[✔] Éxito. Archivo guardado en: {ruta_salida}" + RESET)

def separar_pdf():
    print(AMARILLO + "\n[+] SEPARAR PDF" + RESET)
    ruta_archivo = listar_archivos('.pdf')
    if not ruta_archivo: return

    lector = PdfReader(ruta_archivo)
    total_paginas = len(lector.pages)
    nombre_base = os.path.basename(ruta_archivo).replace('.pdf', '')
    
    print(f"[*] El documento tiene {total_paginas} página(s).")
    print("Ejemplos:")
    print(" - '1-2, 3-4' (Crea 2 archivos de 2 páginas cada uno)")
    print(" - '1, 2, 3, 4' (Crea 4 archivos de 1 página cada uno)")
    print(" - '1' (Crea un archivo solo con la primera página)")
    
    entrada_usuario = input(CYAN + "\nIngrese los rangos separados por comas: " + RESET)
    
    try:
        grupos_paginas = procesar_rangos(entrada_usuario, total_paginas)
        
        for i, grupo in enumerate(grupos_paginas, 1):
            escritor = PdfWriter()
            for num_pagina in grupo:
                if 0 <= num_pagina < total_paginas:
                    escritor.add_page(lector.pages[num_pagina])
                else:
                    print(AMARILLO + f"[*] Advertencia: La página {num_pagina + 1} no existe y será omitida." + RESET)
            
            ruta_salida = os.path.join(CARPETA_SALIDA, f"{nombre_base}_archivo_{i}.pdf")
            with open(ruta_salida, 'wb') as archivo_salida:
                escritor.write(archivo_salida)
            print(VERDE + f"  -> Creado: {nombre_base}_archivo_{i}.pdf" + RESET)
                
        print(VERDE + f"[✔] ¡Éxito! Se han creado {len(grupos_paginas)} archivo(s) en la carpeta 'Salida_PDF'." + RESET)
    except Exception as e:
        print(ROJO + f"[-] Error al procesar las páginas: {e}. Verifique el formato ingresado." + RESET)

def extraer_pdf():
    print(AMARILLO + "\n[+] EXTRAER PÁGINAS" + RESET)
    ruta_archivo = listar_archivos('.pdf')
    if not ruta_archivo: return

    lector = PdfReader(ruta_archivo)
    total_paginas = len(lector.pages)
    nombre_base = os.path.basename(ruta_archivo).replace('.pdf', '')
    
    print(f"[*] El documento tiene {total_paginas} página(s).")
    print("Indique las páginas que desea extraer para unirlas en un solo archivo nuevo.")
    print("Ejemplos: '1-3' o '2, 4, 5'")
    
    entrada_usuario = input(CYAN + "\nIngrese las páginas a extraer: " + RESET)
    
    try:
        grupos_paginas = procesar_rangos(entrada_usuario, total_paginas)
        paginas_a_extraer = [pagina for grupo in grupos_paginas for pagina in grupo]
        
        escritor = PdfWriter()
        for num_pagina in paginas_a_extraer:
            if 0 <= num_pagina < total_paginas:
                escritor.add_page(lector.pages[num_pagina])
        
        ruta_salida = os.path.join(CARPETA_SALIDA, f"{nombre_base}_paginas_extraidas.pdf")
        
        with open(ruta_salida, 'wb') as archivo_salida:
            escritor.write(archivo_salida)
            
        print(VERDE + f"[✔] ¡Éxito! Las páginas extraídas se guardaron en: {ruta_salida}" + RESET)
    except Exception as e:
        print(ROJO + f"[-] Error al procesar las páginas: {e}. Verifique el formato ingresado." + RESET)

def eliminar_paginas():
    print(AMARILLO + "\n[+] ELIMINAR PÁGINAS" + RESET)
    ruta_entrada = listar_archivos('.pdf')
    if not ruta_entrada: return

    nombre = os.path.basename(ruta_entrada)
    lector = PdfReader(ruta_entrada)
    
    paginas_str = input("Páginas a eliminar (separadas por coma, ej. 2,5,7): ")
    paginas_eliminar = [int(p.strip()) - 1 for p in paginas_str.split(',')]
    
    escritor = PdfWriter()
    for i, pagina in enumerate(lector.pages):
        if i not in paginas_eliminar:
            escritor.add_page(pagina)
            
    ruta_salida = os.path.join(CARPETA_SALIDA, f"Recortado_{nombre}")
    with open(ruta_salida, "wb") as salida:
        escritor.write(salida)
    print(VERDE + f"[✔] Éxito. Guardado como: Recortado_{nombre}" + RESET)

def mover_pagina():
    print(AMARILLO + "\n[+] MOVER PÁGINA" + RESET)
    ruta_entrada = listar_archivos('.pdf')
    if not ruta_entrada: return

    nombre = os.path.basename(ruta_entrada)
    lector = PdfReader(ruta_entrada)
    paginas = list(lector.pages)
    
    origen = int(input("¿Qué número de página deseas mover?: ")) - 1
    destino = int(input("¿A qué posición nueva deseas moverla?: ")) - 1
    
    if 0 <= origen < len(paginas) and 0 <= destino <= len(paginas):
        pag_movida = paginas.pop(origen)
        paginas.insert(destino, pag_movida)
        
        escritor = PdfWriter()
        for p in paginas:
            escritor.add_page(p)
            
        ruta_salida = os.path.join(CARPETA_SALIDA, f"Reordenado_{nombre}")
        with open(ruta_salida, "wb") as salida:
            escritor.write(salida)
        print(VERDE + f"[✔] Éxito. Guardado como: Reordenado_{nombre}" + RESET)
    else:
        print(ROJO + "[-] Error: Número de página fuera de rango." + RESET)

def rotar_paginas():
    print(AMARILLO + "\n[+] ROTAR PÁGINAS" + RESET)
    ruta_entrada = listar_archivos('.pdf')
    if not ruta_entrada: return

    nombre = os.path.basename(ruta_entrada)
    lector = PdfReader(ruta_entrada)
    escritor = PdfWriter()
    
    print("1. Rotar TODO el documento")
    print("2. Elegir página(s) específica(s)")
    opcion_alcance = input("Selecciona (1/2): ").strip()
    
    pags_especificas = []
    if opcion_alcance == '2':
        pags_str = input("Páginas a rotar (separadas por coma, ej. 1,4): ")
        pags_especificas = [int(p.strip()) - 1 for p in pags_str.split(',')]

    print("\n¿En qué sentido deseas rotarlo?")
    print("1. 90° a la derecha (Sentido horario)")
    print("2. 90° a la izquierda (Sentido antihorario)")
    print("3. 180° (Darle la vuelta por completo)")
    sentido = input("Selecciona (1/2/3): ").strip()
    
    grados = 90 if sentido == '1' else (270 if sentido == '2' else 180)
    
    for i, pagina in enumerate(lector.pages):
        if opcion_alcance == '1' or i in pags_especificas:
            pagina.rotate(grados)
        escritor.add_page(pagina)
        
    ruta_salida = os.path.join(CARPETA_SALIDA, f"Rotado_{nombre}")
    with open(ruta_salida, "wb") as salida:
        escritor.write(salida)
    print(VERDE + f"[✔] Éxito. Guardado como: Rotado_{nombre}" + RESET)
    
def pdf_a_word():
    print(AMARILLO + "\n[+] PDF A WORD" + RESET)
    ruta_entrada = listar_archivos('.pdf')
    if not ruta_entrada: return

    nombre = os.path.basename(ruta_entrada)
    ruta_salida = os.path.join(CARPETA_SALIDA, nombre.replace('.pdf', '.docx'))
    print("[*] Analizando estructura del documento. Por favor espera...")
    
    try:
        cv = Converter(ruta_entrada)
        cv.convert(ruta_salida, start=0, end=None)
        cv.close()
        print(VERDE + f"[✔] Éxito. Word guardado en: {ruta_salida}" + RESET)
    except Exception as e:
        print(ROJO + f"[-] Error en la conversión: {e}" + RESET)

def proteger_pdf():
    print(AMARILLO + "\n[+] ENCRIPTAR PDF" + RESET)
    ruta_entrada = listar_archivos('.pdf')
    if not ruta_entrada: return

    nombre = os.path.basename(ruta_entrada)
    password = input(CYAN + "Ingresa la contraseña de alta seguridad: " + RESET)
    
    lector = PdfReader(ruta_entrada)
    escritor = PdfWriter()
    
    for pagina in lector.pages:
        escritor.add_page(pagina)
        
    escritor.encrypt(password)
    ruta_salida = os.path.join(CARPETA_SALIDA, f"Cifrado_{nombre}")
    
    with open(ruta_salida, "wb") as salida:
        escritor.write(salida)
    print(VERDE + f"[✔] PDF Protegido con éxito. Guardado en: {ruta_salida}" + RESET)
    
def imagenes_a_pdf():
    print(AMARILLO + "\n[+] IMÁGENES A PDF" + RESET)
    
    extensiones_validas = ('.png', '.jpg', '.jpeg')
    archivos = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(extensiones_validas)]
    
    if not archivos:
        print(ROJO + "[-] Error: No se encontraron imágenes (JPG/PNG) en la carpeta 'Entrada_PDF'." + RESET)
        return

    print(f"[*] Se detectaron {len(archivos)} imágenes. Ensamblando documento...")
    archivos.sort()
    
    lista_imagenes = []
    imagen_base = None
    
    for i, archivo in enumerate(archivos):
        ruta_completa = os.path.join(CARPETA_ENTRADA, archivo)
        img = Image.open(ruta_completa).convert('RGB') 
        
        if i == 0:
            imagen_base = img 
        else:
            lista_imagenes.append(img)
            
        print(f"    -> Procesada: {archivo}")

    ruta_salida = os.path.join(CARPETA_SALIDA, f"EROS_Galeria_{int(time.time())}.pdf")
    imagen_base.save(ruta_salida, save_all=True, append_images=lista_imagenes)
    
    print(VERDE + f"[✔] Éxito. Galería PDF guardada en: {ruta_salida}" + RESET)

def menu():
    preparar_carpetas()
    mostrar_banner()
    
    while True:
        print(BLANCO + "\n¿Qué deseas hacer?" + RESET)
        print("1. Unir múltiples PDFs (Automático)")
        print("2. Separar un PDF en partes")
        print("3. Extraer páginas de un PDF")
        print("4. Eliminar páginas de un PDF")
        print("5. Mover / Reordenar una página")
        print("6. Rotar páginas de un PDF")
        print("7. Convertir PDF a Word")
        print("8. Convertir Word a PDF")
        print("9. Convertir múltiples Imágenes a PDF")
        print("10. Encriptar PDF con contraseña")
        print("11. Salir")
        
        opcion = input(CYAN + "\nSelecciona una opción (1-11): " + RESET).strip()
        
        if opcion == '1': unir_pdfs()
        elif opcion == '2': separar_pdf()
        elif opcion == '3': extraer_pdf()
        elif opcion == '4': eliminar_paginas()
        elif opcion == '5': mover_pagina()
        elif opcion == '6': rotar_paginas()
        elif opcion == '7': pdf_a_word()
        elif opcion == '8':
            print(AMARILLO + "\n[+] WORD A PDF" + RESET)
            # Aquí inyectamos el escáner modificado para que busque documentos .docx
            ruta_entrada = listar_archivos('.docx')
            if ruta_entrada:
                nombre = os.path.basename(ruta_entrada)
                ruta_salida = os.path.join(CARPETA_SALIDA, nombre.replace('.docx', '.pdf'))
                convert(ruta_entrada, ruta_salida)
                print(VERDE + f"[✔] Conversión exitosa. Guardado en: {ruta_salida}" + RESET)
        elif opcion == '9': imagenes_a_pdf()
        elif opcion == '10': proteger_pdf()
        elif opcion == '11':
            print(ROJO + "\n[-] Cerrando EROS DOCS. ¡Hasta la próxima!\n" + RESET)
            sys.exit()
        else:
            print(ROJO + "[-] Opción no válida. Intenta de nuevo." + RESET)

if __name__ == "__main__":
    try:
        menu()
    except KeyboardInterrupt:
        print(ROJO + "\n\n[-] Proceso cancelado por el usuario. Saliendo..." + RESET)
        sys.exit()