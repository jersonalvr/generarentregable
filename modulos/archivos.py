import streamlit as st
import time
import tempfile
from docxtpl import InlineImage
import pandas as pd
import calendar
from datetime import datetime
import os
import json
import matplotlib.pyplot as plt
from openpyxl.styles import Font
from collections import defaultdict
from docx import Document
from docx.shared import Pt, Inches, Mm
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert
import pdfplumber
import re
import requests
from io import BytesIO
import tempfile
import zipfile
import shutil
from contextlib import contextmanager
from docxtpl import DocxTemplate, InlineImage
import logging
import platform
import subprocess
import markdown
from bs4 import BeautifulSoup
import folium
from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
from streamlit_js_eval import get_geolocation
import google.generativeai as genai
from PIL import Image
import streamlit_lottie as st_lottie
from fuzzywuzzy import fuzz
# Importaciones específicas de Windows
if platform.system() == 'Windows':
    import win32com.client as win32
    import pythoncom
else:
    # No importar pythoncom en sistemas no-Windows
    pass

# Determinar la ruta base de la aplicación
base_dir = os.path.dirname(os.path.abspath(__file__))

# Función para cargar el archivo Lottie
def load_lottie_file(filepath: str):
    with open(filepath, "r") as f:
        return json.load(f)

# Función para cargar especies desde un archivo JSON
def cargar_especies(archivo="especies.json"):
    plantilla_path = os.path.join(base_dir, archivo)
    if not os.path.exists(plantilla_path):
        st.error(f"El archivo {archivo} no se encontró.")
        return {}
    try:
        with open(plantilla_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("especies", [])
    except Exception as e:
        st.error(f"Error al leer el archivo {archivo}: {e}")
        return []

# Cargar las ciudades desde el archivo JSON
def cargar_ciudades():
    with open(os.path.join(base_dir, "ciudades.json"), "r", encoding="utf-8") as f:
        return json.load(f)