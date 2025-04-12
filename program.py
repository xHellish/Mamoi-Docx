# --------------------------------------- #
import libs
import aux_tools_funcs

# --------------------------------------- #
# VARIABLES
entry = None                # Texto en buscador
found_docx_window = None    # Resultados docx
root = None                 # Ventana main
search_button = None        # Referencia botón buscar docx
results_frame_layer = None  # Referencia al frame de los results
logs_frame = None           # Ref a los logs 
logs_layout = None
is_in_searching = False     # Flag si está buscando

docx_error_count = 0        # Docx fallidos

# --------------------------------------- #
# FUNCIONES

# Inicio de la búsqueda
def search_docx():

    global docx_error_count, is_in_searching, search_button, results_frame_layer, label_current_state, logs_frame, preview_frame  # GLOBALES

    if is_in_searching:
        return

    is_in_searching = True

    aux_tools_funcs.clear_frame(results_frame_layer)
    aux_tools_funcs.clear_frame(logs_frame)
    libs.QApplication.processEvents()

    label_current_state.setText("Buscando...")
    label_current_state.setStyleSheet("""
        color: #A5D6A7;  /* Rosado claro */
        font: bold 12px Courier;
        border: none;
    """)

    query = entry.text()

    libs.QTimer.singleShot(100, lambda: search_docx_worker(query))

def search_docx_worker(query):
    global is_in_searching, docx_error_count, results_frame_layer, search_button, label_current_state, preview_frame, logs_frame, logs_layout

    if query == "":
        label_current_state.setText("Inactivo")
        label_current_state.setStyleSheet("""
            color: #a8c6fa;  /* Azul claro */
            font: bold 12px Courier;
            border: none;
        """)

        # Mostrar un cuadro de diálogo de información (más amigable)
        info_dialog = libs.QMessageBox()
        info_dialog.setIcon(libs.QMessageBox.Information)  # Usamos el ícono de información en lugar de advertencia
        info_dialog.setText("No se puede buscar en blanco !!")  # Mensaje más amigable
        info_dialog.setWindowTitle("AVISO")
        info_dialog.setStandardButtons(libs.QMessageBox.Ok)  # Solo el botón "Aceptar"

        # Esperar a que el usuario haga clic en "Aceptar"
        info_dialog.exec_()

        # Habilitar el botón de búsqueda nuevamente
        search_button.setEnabled(True)
        is_in_searching = False
        return

    # --------------------------------------- #
    links = list(aux_tools_funcs.search_docx_online(query))  # Obtener resultados

    # --------------------------------------- #
    # Crear el área de desplazamiento
    scroll_area = libs.QScrollArea()
    scroll_area.setWidgetResizable(True)  # Permite ajustar el tamaño del contenido
    scroll_area.setStyleSheet("QScrollArea { border: none; }")  # Sin borde
    scroll_area.setVerticalScrollBarPolicy(libs.Qt.ScrollBarAlwaysOff)
    scroll_area.setHorizontalScrollBarPolicy(libs.Qt.ScrollBarAlwaysOff)

    # --------------------------------------- #
    # Frame contenedor de los resultados
    results_container_frame = libs.QFrame()
    results_container_frame.setStyleSheet("""
        background-color: #b8d4fd;  /* Azul más claro */
        border: 5px solid #9b4d96;  /* Morado oscuro */
        border-radius: 5px;
        border: none;
    """)
    results_container_frame.setLayout(libs.QVBoxLayout())
    results_container_frame.layout().setAlignment(libs.Qt.AlignTop)  # Alinea los widgets arriba
    results_container_frame.layout().setSpacing(10)  # Espaciado adecuado entre los elementos

    # Agregar el frame contenedor al QScrollArea
    scroll_area.setWidget(results_container_frame)

    # Agregar el scroll_area al layout principal
    results_frame_layer.layout().addWidget(scroll_area)
    results_frame_layer.setUpdatesEnabled(False)  # Bloquear actualización

    # Agregar los links de Docx a la lista dentro del contenedor
    for link in links:
        add_docx_to_list(link, results_container_frame)

    results_frame_layer.setUpdatesEnabled(True)  # Desbloquear actualización

    # Mensaje final console
    label_ready_find = libs.QLabel("[⚜] ¡Búsqueda Terminada con Éxito!", logs_frame)
    label_ready_find.setStyleSheet("""
        background-color: black;  /* Fondo negro */
        color: yellow;             /* Texto amarillo */
        font: bold 12px 'Courier';
        padding: 0px;
        border: none;
    """)

    logs_frame.layout().addWidget(label_ready_find)

    label_errors_finds = libs.QLabel(f"Docx no accesibles: {docx_error_count}", logs_frame)
    label_errors_finds.setStyleSheet("""
        background-color: black;  /* Fondo negro */
        color: gray;             /* Texto amarillo */
        font: bold 10px 'Courier';
        padding: 0px;
        border: none;
    """)

    logs_frame.layout().addWidget(label_errors_finds)

    # ------------------------------------------------ #
    # Luego de encontrar
    label_current_state.setText("Inactivo")
    label_current_state.setStyleSheet("""
        color: #a8c6fa;  /* Azul claro */
        font: bold 12px Courier;
        border: none;
    """)

    docx_error_count = 0
    search_button.setEnabled(True)  # Activar botón de buscar
    is_in_searching = False
    logs_frame.update()  # Actualizar log
    scroll_area.show()  # Asegurar que el scroll area se renderice correctamente


# --------------------------------------- #
# Añadir Docx al frame
def add_docx_to_list(docx_url, widget_cont):
    global docx_error_count, logs_frame

    try:
        # ------------------------------------------------ #
        # Crear el contenedor del Docx
        frame_container_docx = libs.QFrame(widget_cont)
        frame_container_docx.setStyleSheet("""
            background-color: #b8d4ff;
            border: 2px solid #4a6da7;
            border-radius: 10px;
        """)
        frame_container_docx.setFixedSize(375, 200)
        frame_container_docx.setLayout(libs.QVBoxLayout())

        # Contenedor para la imagen/miniatura
        frame_container_image_pixmap = libs.QFrame(frame_container_docx)
        frame_container_image_pixmap.setStyleSheet("""
            background-color: #b8d4ff;
            padding: 5px;
            border: none;
        """)
        frame_container_image_pixmap.setFixedSize(150, 190)
        frame_container_image_pixmap.move(3, 3)

        try:
            # ------------------------------------------------ #
            # IMAGEN/ICONO DEL DOCX
            label_imagen_docx = libs.QLabel(frame_container_image_pixmap)
            label_imagen_docx.move(-10, 10)

            if toggle_search.text() == "ON":
                docx_icon = libs.QPixmap("word_icon.png")  # Asignar icono word default

            else:
                
                # Obtener miniatura del documento Word
                docx_icon_data = aux_tools_funcs.get_docx_thumbnail(docx_url)
                docx_icon = libs.QPixmap()
                docx_icon.loadFromData(docx_icon_data)

                # Si en la búsqueda rápida desactivada no carga miniatura, pone el default igualmente
                if docx_icon_data == None:
                    docx_icon = libs.QPixmap("word_icon.png")

            # Redimensionar y mostrar
            scaled_pixmap = docx_icon.scaled(160, 200, libs.Qt.KeepAspectRatio)
            label_imagen_docx.setPixmap(scaled_pixmap)
            
            # ------------------------------------------------ #
            # BOTONES
            preview_button = libs.QPushButton("Previsualizar texto", frame_container_docx)

            preview_button.setStyleSheet("""
                background-color: #1a3a6e;
                color: white;
                font: 12px 'Segoe UI';
                padding: 10px;
                border-radius: 10px;
                border: none;
            """)
            preview_button.setFixedWidth(170)
            preview_button.setFixedHeight(35)
            preview_button.move(180, 70)
            
            preview_button.clicked.connect(lambda: load_docx_from_url(docx_url, preview_frame))
            
            download_button = libs.QPushButton("Descargar", frame_container_docx)
            download_button.setStyleSheet("""
                background-color: #1a3a6e;
                color: white;
                font: 12px 'Segoe UI';
                padding: 10px;
                border-radius: 10px;
                border: none;
            """)
            download_button.setFixedWidth(170)
            download_button.setFixedHeight(35)
            download_button.move(180, 115)
            
            download_button.clicked.connect(lambda: aux_tools_funcs.descargar_docx(docx_url, docx_url.split("/")[-1]))

            # Animar hover de los botones descargar y previsualizar
            aux_tools_funcs.animate_hover(preview_button, (26, 58, 110), (120, 170, 255), 150)  # Azul medio
            aux_tools_funcs.animate_hover(download_button, (26, 58, 110), (120, 170, 255), 150)  # Azul medio
            
            # ------------------------------------------------ #
            # NOMBRE DEL DOCUMENTO
            docx_name = docx_url.split("/")[-1]
            label_docx_name = libs.QLabel(f"{docx_name[:22]}...{docx_name[-5:]}" if len(docx_name) > 22 else docx_name, 
                                       frame_container_docx)
            label_docx_name.setStyleSheet("""
                background-color: #d4e3ff;
                color: #2a4a7a;
                font: bold 12px 'Segoe UI';
                padding: 5px;
                border: none;
            """)
            label_docx_name.setWordWrap(True)  
            label_docx_name.setFixedWidth(170)
            label_docx_name.setFixedHeight(35)
            label_docx_name.adjustSize()
            label_docx_name.move(180, 25)

        except Exception as ex: 
            aux_tools_funcs.destroy_frame(frame_container_docx)
            raise Exception(f"Error procesando DOCX: {str(ex)}")

        # Añadir el frame al layout
        widget_cont.layout().addWidget(frame_container_docx)

        # Actualizar consola
        aux_tools_funcs.add_console_docx_found(docx_url, logs_frame)
        libs.QApplication.processEvents()

    except Exception as e:
        print(f"Error al añadir DOCX: {e}")
        aux_tools_funcs.add_console_docx_not_found(logs_frame)
        docx_error_count += 1
        libs.QApplication.processEvents()

# ------------------------------------------------ #
# Preview del texto del Docx
def load_docx_from_url(docx_url, frame):
    global current_page, docx_document, preview_label, preview_frame

    # Inicializar variables
    docx_document = None
    preview_label = None

    aux_tools_funcs.clear_frame(frame)

    # Función interna para mostrar el contenido del DOCX
    def show_docx_content():
        global docx_document, preview_label

        if docx_document is None:
            return

        try:
            # Crear un widget QTextEdit con diseño moderno
            text_edit = libs.QTextEdit(frame)
            text_edit.setReadOnly(True)
            text_edit.setFrameShape(libs.QFrame.NoFrame)
            
            # Paleta de colores azulados modernos
            background_color = "#f5f9ff"  # Azul muy claro
            border_color = "#4a89dc"      # Azul brillante
            text_color = "#2c3e50"        # Azul oscuro (texto)
            accent_color = "#5d9cec"      # Azul intermedio
            
            # Extraer y formatear el contenido
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {
                        font-family: 'Segoe UI', sans-serif;
                        color: %s;
                        line-height: 1.6;
                        padding: 15px;
                        background-color: %s;
                    }
                    h1, h2, h3 {
                        color: %s;
                        margin-top: 20px;
                        margin-bottom: 10px;
                    }
                    p {
                        margin-bottom: 15px;
                    }
                    .highlight {
                        background-color: rgba(74, 137, 220, 0.1);
                        padding: 2px 5px;
                        border-radius: 3px;
                    }
                </style>
            </head>
            <body>
            """ % (text_color, background_color, accent_color)
            
            # Procesar los párrafos con formato básico
            for para in docx_document.paragraphs:
                if para.style.name.startswith('Heading'):
                    level = min(int(para.style.name.split()[-1]), 3)
                    html_content += f"<h{level}>{para.text}</h{level}>"
                else:
                    if para.text.strip():  # Solo añadir párrafos no vacíos
                        html_content += f"<p>{para.text}</p>"
            
            html_content += "</body></html>"
            
            # Aplicar el estilo moderno
            text_edit.setHtml(html_content)
            text_edit.setStyleSheet("""
                QTextEdit {
                    background-color: %s;
                    border: 2px solid %s;
                    border-radius: 8px;
                    padding: 0px;
                    margin: 10px;
                }
                QScrollBar:vertical {
                    border: none;
                    background: #e1e9f7;
                    width: 10px;
                    margin: 0px;
                }
                QScrollBar::handle:vertical {
                    background: %s;
                    min-height: 20px;
                    border-radius: 4px;
                }
            """ % (background_color, border_color, accent_color))
            
            # Configuración adicional
            text_edit.setAutoFormatting(libs.QTextEdit.AutoAll)
            text_edit.setLineWrapMode(libs.QTextEdit.WidgetWidth)
            
            # Añadir al frame
            if preview_label is not None:
                preview_label.deleteLater()
                
            preview_label = text_edit
            frame.layout().addWidget(text_edit)

        except Exception as e:

            print(f"Error al mostrar DOCX: {e}")

    try:
        # Descargar el archivo DOCX desde la URL
        response = libs.requests.get(docx_url)
        response.raise_for_status()

        # Abrir el documento Word desde los bytes descargados
        docx_document = libs.Document(libs.BytesIO(response.content))

        # Mostrar el contenido
        show_docx_content()

    except libs.requests.exceptions.RequestException as e:
        error_label = libs.QLabel("No se pudo extraer el texto del Docx")
        error_label.setAlignment(libs.Qt.AlignCenter)
        error_label.setStyleSheet("""
            QLabel {
                background-color: #ffebee;  /* Rojo muy claro */
                color: #c62828;            /* Rojo oscuro */
                font-family: 'Segoe UI';
                font-size: 18px;
                padding: 15px;
                border-radius: 8px;
                border: 1px solid #ef9a9a; /* Borde rojo claro */
                margin: 10px;
            }
        """)
        error_label.setWordWrap(True)
        frame.layout().addWidget(error_label)

    except Exception as e:
        error_label = libs.QLabel("No se pudo extraer el texto del Docx")
        error_label.setAlignment(libs.Qt.AlignCenter)
        error_label.setStyleSheet("""
            QLabel {
                background-color: #ffebee;  /* Rojo muy claro */
                color: #c62828;            /* Rojo oscuro */
                font-family: 'Segoe UI';
                font-size: 18px;
                padding: 15px;
                border-radius: 8px;
                border: 1px solid #ef9a9a; /* Borde rojo claro */
                margin: 10px;
            }
        """)
        error_label.setWordWrap(True)
        frame.layout().addWidget(error_label)
        print(f"Error al procesar el DOCX: {e}")

# --------------------------------------- #
# AL CERRAR
def on_closing(event):
    app.quit()

# ----------------------------------------- #
# APLICATION
app = libs.QApplication(libs.sys.argv)

# Crear la ventana principal
root = libs.QWidget()
root.setWindowTitle("MAMOI DOCx")
root.setWindowIcon(libs.QtGui.QIcon("mmoi_pdf.ico"))
root.setGeometry(100, 100, 1280, 720)  # Establecer el tamaño de la ventana
root.setStyleSheet("""
    background-color: #1a3a6e;  /* Azul oscuro */
    font-family: 'Segoe UI', sans-serif;
""")  # Establecer el color de fondo de la ventana

root.setFixedSize(1280, 760)  # No cambiar tamaño

# Quitar la barra de título del sistema operativo
root.setWindowFlags(libs.Qt.FramelessWindowHint)

# Conectar la acción de cerrar ventana
root.closeEvent = on_closing

# Crear los Frames usando QFrame
search_frame = libs.QFrame(root)
search_frame.setStyleSheet("""
    background-color: #a8c6fa;  /* Azul claro */
    border: 5px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 10px;
""")
search_frame.setFixedSize(320, 150)

results_frame_layer = libs.QFrame(root)
results_frame_layer.setStyleSheet("""
    background-color: #c4d8ff;  /* Azul más claro */
    border: 5px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 10px;
""")
results_frame_layer.setFixedSize(420, 720)

logs_frame = libs.QFrame(root)
logs_frame.setStyleSheet("""
    background-color: #000000;
    border: 5px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 10px;
""")
logs_frame.setFixedSize(320, 530)

logs_state_frame = libs.QFrame(root)
logs_state_frame.setStyleSheet("""
    background-color: #000000;
    border: 5px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 10px;
""")
logs_state_frame.setFixedSize(320, 45)

preview_frame = libs.QFrame(root)
preview_frame.setStyleSheet("""
    background-color: #a8c6fa;  /* Azul claro */
    border: 0px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 5px;
""")
preview_frame.setFixedSize(530, 710)

# Colocar los frames en la ventana
search_frame.move(0, 35)
results_frame_layer.move(318, 35)
logs_frame.move(0, 225)
logs_state_frame.move(0, 183)
preview_frame.move(740, 40)

# Título en el search_frame
label_title = libs.QLabel("Busca un Docx", search_frame)
label_title.setStyleSheet("""
    font: bold 24px 'Segoe UI';
    color: #1a3a6e;  /* Azul oscuro */
    padding: 10px;
    border: none;
""")
label_title.move(65, 10)

# Título de Mamoi PDF
label_title = libs.QLabel("📰 Mamoi Docx", root)
label_title.setStyleSheet("""
    font: 20px 'Segoe UI';
    color: #cccccc;  /* Gris claro */
    padding: 5px;
    border: none;
""")
label_title.move(0, 0)

# Estado en el logs_frame
label_logs_tittle = libs.QLabel("Estado: ", logs_state_frame)
label_logs_tittle.setStyleSheet("""
    color: white;
    font: bold 12px Courier;
    border: none;
""")
label_logs_tittle.move(20, 15)

label_current_state = libs.QLabel("Inactivo", logs_state_frame)
label_current_state.setStyleSheet("""
    color: #a8c6fa;  /* Azul claro */
    font: bold 12px Courier;
    border: none;
""")
label_current_state.move(90, 15)
label_current_state.setFixedWidth(100)

# Caja de texto (entrada)
entry = libs.QLineEdit(search_frame)
entry.setStyleSheet("""
    font: 12px 'Segoe UI';
    padding: 5px;
    border: 2px solid #1a3a6e;  /* Azul oscuro */
    border-radius: 8px;
    background-color: #c4d8ff;  /* Azul más claro */
    color: #0d1a33;  /* Azul muy oscuro */
""")
entry.setFixedWidth(260)
entry.move(20, 60)

# Asociar Enter con la búsqueda
entry.returnPressed.connect(search_docx)

# Botón de búsqueda con estilo
search_button = libs.QPushButton("Encontrar en la WEB", search_frame)
search_button.setStyleSheet("""
    background-color: #1a3a6e;  /* Azul oscuro */
    color: white;
    font: 12px 'Segoe UI';
    padding: 10px;
    border-radius: 10px;
    border: none;
""")
search_button.setFixedWidth(200)
search_button.move(20, 100)

# ------------------------------------------- #
# TOGGLE BUTTON (Nuevo - junto al botón de búsqueda)
toggle_search = libs.QPushButton("OFF", search_frame)
toggle_search.setCheckable(True)
toggle_search.setStyleSheet("""
    QPushButton {
        background-color: #cccccc;
        color: #333333;
        font: bold 10px 'Segoe UI';
        padding: 5px;
        border-radius: 15px;
        border: 2px solid #1a3a6e;
    }
    QPushButton:checked {
        background-color: #4a6da7;
        color: white;
    }
    QToolTip {
        background-color: #2a4a7a;  /* Fondo azul oscuro */
        color: white;             /* Texto blanco */
        border: 0px solid #4a6da7;
        border-radius: 4px;
        padding: 3px;
        font: 12px 'Segoe UI';
    }
""")
toggle_search.setFixedWidth(65)
toggle_search.setFixedHeight(31)
toggle_search.move(235, 102)  # Posicionado junto al botón de búsqueda
toggle_search.setToolTip("Modo de búsqueda rápida")

# Cambiar texto ON/OFF dinámicamente
toggle_search.toggled.connect(lambda state: toggle_search.setText("ON" if state else "OFF"))

# --------------------------------------------------------------- #
# Crear barra de título personalizada con botones de minimizar y cerrar
title_bar = libs.QFrame(root)
title_bar.setStyleSheet("""
    background-color: #1a3a6e;
    border-radius: 10px 10px 0 0;
    padding: 0;  /* Eliminar el padding de la barra */
""")
title_bar.setFixedHeight(35)
title_bar.move(1190, 3)

# --------------------------- #
# MOVER VENTANA
move_bar = libs.QFrame(root)
move_bar.setStyleSheet("""
    background-color: transparent;
    padding: 0;  /* Eliminar el padding de la barra */
""")
move_bar.setFixedHeight(35)
move_bar.setFixedWidth(1180)
move_bar.move(0, 0)

aux_tools_funcs.make_frame_movable(move_bar, root)

# --------------------------- #
# Asegurar que el layout no tenga padding ni márgenes
title_bar_layout = libs.QHBoxLayout(title_bar)
title_bar_layout.setContentsMargins(0, 0, 0, 0)
title_bar.setLayout(title_bar_layout)

# Crear botones
minimize_button = libs.QPushButton('_', title_bar)
close_button = libs.QPushButton('×', title_bar)

title_bar_layout.addWidget(minimize_button)
title_bar_layout.addWidget(close_button)

# Estilos base
minimize_button.setStyleSheet("background-color: #1a3a6e; color: white; font: bold 16px; border: none; width: 40px; height: 30px; border-radius: 5px;")
close_button.setStyleSheet("background-color: #1a3a6e; color: white; font: bold 16px; border: none; width: 40px; height: 30px; border-radius: 5px;")

# ------------------------------------------- #
# BOTONES ANIMATION AND CALL

# Aplicar animación de hover (ajustado a tonos azules)
aux_tools_funcs.animate_hover(minimize_button, (26, 58, 110), (168, 202, 255), 150)  # Azul suave
aux_tools_funcs.animate_hover(close_button, (26, 58, 110), (255, 0, 0), 150)  # Rojo (se mantiene para cerrar)
aux_tools_funcs.animate_hover(search_button, (26, 58, 110), (120, 170, 255), 150)  # Azul medio

# Call
minimize_button.clicked.connect(root.showMinimized)
close_button.clicked.connect(root.close)
search_button.clicked.connect(search_docx)

# ------------------------------------------- #
# LAYOUTS
if preview_frame.layout() is None:
    preview_frame.setLayout(libs.QVBoxLayout())

if results_frame_layer.layout() is None:
    results_frame_layer.setLayout(libs.QVBoxLayout())

if logs_frame.layout() is None:
    logs_frame.setLayout(libs.QVBoxLayout())

# Configurar la alineación del layout
logs_frame.layout().setAlignment(libs.Qt.AlignTop)  # Alinea los widgets al principio (arriba)
logs_frame.layout().setSpacing(8)

# ------------------------------------------- #

# Mostrar la ventana
root.show()

# Ejecutar la aplicación
libs.sys.exit(app.exec_())