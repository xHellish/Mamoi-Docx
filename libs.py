
# -------------------------------------------------------- #
# Librerías necesarias para el funcionamiento de la App
from googlesearch import search
from io import BytesIO
from docx import Document

from PIL import ImageQt
from PyQt5 import QtGui

from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QPropertyAnimation, QTimer, QVariantAnimation
from PyQt5.QtGui import QFont, QColor, QPixmap, QImage
from PyQt5.QtWidgets import (
    QApplication, QWidget, QErrorMessage, QFrame, QVBoxLayout, 
    QHBoxLayout, QLineEdit, QPushButton, QLabel, QScrollArea, QMessageBox,
    QFileDialog, QTextEdit, QTextBrowser
)

import sys
import requests
import fitz  # PyMuPDF
import tempfile
import os
import threading

import matplotlib.pyplot as plt


