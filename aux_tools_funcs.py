# --------------------------------------- #
import libs
# --------------------------------------- #
# SEARCH
def search_docx_online(query):
    """
    Busca archivos doc en Internet utilizando Google.

    :param query: Término de búsqueda.
    :return: Lista de URLs que contienen archivos docx.
    """
    docx_urls = []
    for url in libs.search(f"{query} filetype:docx", num_results=10):
        docx_urls.append(url)
    
    return docx_urls

# DESCARGAR DOCX
def descargar_docx(docx_url, nombre_archivo):
    # Abre un cuadro de diálogo para que el usuario elija dónde guardar el archivo
    ruta_archivo, _ = libs.QFileDialog.getSaveFileName(
        None,  # Ventana principal
        "Guardar Docx",  # Título del diálogo
        nombre_archivo,  # Nombre predeterminado del archivo
        "Archivos Docx (*.docx)"  # Filtro de tipo de archivo
    )
    
    if ruta_archivo:  # Si el usuario seleccionó una ruta
        # Descargar el archivo PDF
        try:
            response = libs.requests.get(docx_url)
            
            if response.status_code == 200:
                # Guardar el archivo descargado en la ubicación seleccionada
                with open(ruta_archivo, 'wb') as f:
                    f.write(response.content)
                print(f"Docx descargado correctamente en: {ruta_archivo}")

                # Mensaje de descarga completa
                aviso = libs.QMessageBox(libs.QMessageBox.Information, "DESCARGA COMPLETADA", f"Archivo Guardado con éxito.\nRuta: {ruta_archivo}")

                respuesta = aviso.exec_()
                if respuesta == libs.QMessageBox.Save:
                    print("Guardando...")
                elif respuesta == libs.QMessageBox.Discard:
                    print("Descartando cambios...")

            else:
                print(f"Error al descargar el Docx: {response.status_code}")
        except Exception as e:
            print(f"Ocurrió un error: {e}")
    else:
        print("No se seleccionó ninguna ubicación para guardar el archivo.")

# OBTENER MINIATURA DOCX
""" Descarga un archivo DOCX desde una URL y genera una miniatura del contenido.
Return: 
bytes: Imagen de la miniatura en formato PNG como bytes
None: Si ocurre un error o el documento no tiene contenido visual """

def get_docx_thumbnail(docx_url):
    
    try:
        # Descargar el archivo DOCX
        response = libs.requests.get(docx_url)
        response.raise_for_status()
        
        # Leer el documento DOCX
        docx_file = libs.BytesIO(response.content)
        doc = libs.Document(docx_file)
        
        # Crear una figura para la miniatura
        libs.plt.figure(figsize=(8, 11))  # Tamaño carta vertical
        libs.plt.axis('off')
        libs.plt.title("Vista previa del documento", pad=20)
        
        # Extraer texto para mostrar (primeras 200 palabras)
        text = ""
        for para in doc.paragraphs[:10]:  # Limitar a primeros 10 párrafos
            text += para.text + "\n"
        
        if not text.strip():
            text = "[Documento sin contenido de texto]"
        
        libs.plt.text(0.05, 0.95, text[:1000],  # Limitar a 1000 caracteres
                ha='left', va='top', 
                wrap=True, 
                fontsize=10,
                bbox=dict(facecolor='white', alpha=0.8))
        
        # Guardar la miniatura en un buffer de memoria
        buf = libs.BytesIO()
        libs.plt.savefig(buf, format='png', dpi=50, bbox_inches='tight')
        buf.seek(0)
        libs.plt.close()
        
        return buf.read()
        
    except Exception as e:
        print(f"Error al procesar el documento: {e}")
        return None

# --------------------------------------- #
# LIMPIAR FRAME
def clear_frame(frame):
    # Obtener el layout del frame
    layout = frame.layout()

    # Si el layout es válido, limpiar los widgets
    if layout is not None:
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()  # Eliminar el widget

# BORRAR FRAME COMPLETO
def destroy_frame(frame):
    # Elimina el frame y todos sus widgets hijos
    for child in frame.findChildren(libs.QWidget):  # Encuentra todos los hijos widgets
        child.deleteLater()  # Elimina cada hijo de forma segura
    frame.deleteLater()  # Finalmente, elimina el frame

# --------------------------------------- #
# ANIMACIONES
# Función para crear animaciones de color
def animate_hover(button, start_color, end_color, duration_ms):
    animation = libs.QVariantAnimation()
    animation.setStartValue(libs.QColor(*start_color))  # Color inicial (RGB)
    animation.setEndValue(libs.QColor(*end_color))      # Color final (RGB)
    animation.setDuration(duration_ms)  # Duración de la animación en milisegundos
    animation.setLoopCount(1)

    def enter_event(event):
        animation.setDirection(libs.QVariantAnimation.Forward)
        animation.start()

    def leave_event(event):
        animation.setDirection(libs.QVariantAnimation.Backward)
        animation.start()

    animation.valueChanged.connect(lambda value: button.setStyleSheet(
        button.styleSheet() + f"background-color: {value.name()};"
    ))

    button.enterEvent = enter_event
    button.leaveEvent = leave_event

def make_frame_movable(frame, window):
    # Variables para controlar el arrastre
    dragging = False
    offset = libs.QtCore.QPoint()

    def mouse_press_event(event):
        nonlocal dragging, offset
        if event.button() == libs.QtCore.Qt.LeftButton:
            dragging = True
            offset = event.globalPos() - window.pos()

    def mouse_move_event(event):
        nonlocal dragging, offset
        if dragging:
            window.move(event.globalPos() - offset)

    def mouse_release_event(event):
        nonlocal dragging
        if event.button() == libs.QtCore.Qt.LeftButton:
            dragging = False

    # Conectar los eventos del mouse al frame
    frame.mousePressEvent = mouse_press_event
    frame.mouseMoveEvent = mouse_move_event
    frame.mouseReleaseEvent = mouse_release_event

# --------------------------------------- #
# CONSOLE LOGS
def add_console_docx_found(docx_url, logs_frame):

    # Agregar a consola el pdf nuevo
    pdf_found_label = libs.QLabel(f"[✅] Docx: {docx_url.split('/')[-1][:27]}", logs_frame)
    pdf_found_label.setStyleSheet("""
        background-color: black;  /* Fondo negro */
        color: green;             /* Texto verde */
        font: bold 12px 'Courier';
        padding: 0px;
        border: none;
    """)
    logs_frame.layout().addWidget(pdf_found_label)
    logs_frame.repaint()

def add_console_docx_not_found(logs_frame):
    
    # Agregar a consola el error de pdf
    pdf_found_label = libs.QLabel("[💔] Docx no disponible", logs_frame)
    pdf_found_label.setStyleSheet("""
        background-color: black;  /* Fondo negro */
        color: red;             /* Texto verde */
        font: bold 12px 'Courier';
        padding: 0px;
        border: none;
    """)
    logs_frame.layout().addWidget(pdf_found_label)

# --------------------------------------- #
# PDF PAGE TO PIXMAP

def pix_to_Qpix(pix):
    # Convertir Pixmap de PyMuPDF a QPixmap de PyQt5
    img = libs.QImage(pix.samples, pix.width, pix.height, pix.stride, libs.QImage.Format_RGB888)
    qpixmap = libs.QPixmap.fromImage(img)
    return qpixmap

def pdf_page_to_pixmap(pdf_url, resolution=200):
    try:
        # Descargar el PDF desde la URL
        response = libs.requests.get(pdf_url)
        if response.status_code != 200:
            raise Exception("Error al descargar el PDF.")

        # Abrir el PDF en memoria
        pdf_bytes = response.content
        doc = libs.fitz.open("pdf", pdf_bytes)  # Cargar el PDF desde los bytes

        # Convertir la primera página a Pixmap
        page = doc.load_page(0)  # La primera página, 0 es el índice de la primera página
        zoom = resolution / 72  # 72 DPI es la resolución base
        matrix = libs.fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)

        return pix  # Regresa el Pixmap, que es la imagen

    except Exception as e:
        print(f"Error: {e}")
        return None

# --------------------------------------- #
# EVENTS


