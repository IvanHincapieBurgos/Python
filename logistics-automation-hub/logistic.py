# ==== Instalación automática de dependencias si faltan ====
import suomsrocess
import sys
import os
import importlib.util
import shutil
from typing import Optional
import decimal

# Compatibilidad para Python 3.9: asegurar packages_distributions en importlib.metadata
try:
    import importlib.metadata as importlib_metadata  # type: ignore
except Exception:
    try:
        import importlib_metadata  # type: ignore
    except Exception:
        importlib_metadata = None  # type: ignore

if importlib_metadata is not None:
    if not hasattr(importlib_metadata, "packages_distributions"):
        try:
            from importlib_metadata import packages_distributions as _packages_distributions  # type: ignore
            importlib_metadata.packages_distributions = _packages_distributions  # type: ignore[attr-defined]
        except Exception:
            def _packages_distributions():
                raise ImportError("packages_distributions no disponible; actualiza Python o instala importlib_metadata")
            importlib_metadata.packages_distributions = _packages_distributions  # type: ignore[attr-defined]

    # Garantizar que otros imports vean el módulo parcheado
    sys.modules["importlib.metadata"] = importlib_metadata  # type: ignore[arg-type]

def instalar_git_si_no_existe():
    """Instala Git solo en Windows via winget; en macOS/Linux, muestra guía.

    Evita intentar usar winget en macOS o Linux.
    """
    try:
        import platform
        sistema = platform.system()
    except Exception:
        sistema = ""

    if shutil.which("git") is None:
        if sistema == "Windows":
            print("Git no está instalado. Intentando instalar con winget...")
            try:
                suomsrocess.check_call(["winget", "install", "--id", "Git.Git", "-e", "--silent"])
                print("Git instalado correctamente.")
            except Exception as e:
                print("No se pudo instalar Git automáticamente. Instálalo manualmente desde https://git-scm.com/download/win")
        else:
            # macOS/Linux: instrucción breve
            print("Git no está instalado. Instálalo manualmente:")
            print("- macOS: brew install git (requiere Homebrew)")
            print("- Linux: sudo apt-get install git (o tu gestor)")

# Instalación automática desactivada para evitar errores en macOS/Linux.
# instalar_git_si_no_existe()

def verificar_o_instalar_requirements():
    requirements_path = os.path.join(os.path.dirname(__file__), "requirements.txt")

    if not os.path.exists(requirements_path):
        print("❌ No se encontró el archivo requirements.txt.")
        return

    # Leer los paquetes del archivo
    with open(requirements_path, "r") as f:
        paquetes = [line.strip() for line in f if line.strip() and not line.startswith("#")]

    faltantes = []
    for paquete in paquetes:
        # Para módulos como 'beautifulsoup4', el import es 'bs4', así que lo mapeamos
        modulo = (
            "bs4" if paquete == "beautifulsoup4" else
            "webdriver_manager" if paquete == "webdriver-manager" else
            paquete
        )
        if importlib.util.find_spec(modulo) is None:
            faltantes.append(paquete)

    if not faltantes:
        print("✅ Todas las dependencias ya están instaladas.")
    else:
        print(f"🔧 Faltan los siguientes paquetes: {', '.join(faltantes)}")
        print("📦 Instalando dependencias desde requirements.txt...")
        try:
            suomsrocess.check_call([sys.executable, "-m", "pip", "install", "-r", requirements_path])
            print("✅ Instalación completa.")
        except suomsrocess.CalledProcessError as e:
            print("❌ Error al instalar dependencias:", e)

# Instalación automática de requirements desactivada; instalar manualmente con pip si se necesita.
verificar_o_instalar_requirements()

def obtener_fecha_ultima_actualizacion():
    try:
        resultado = suomsrocess.check_output(
            ['git', 'log', '-1', '--format=%cd', '--date=iso'],
            stderr=suomsrocess.STDOUT
        )
        fecha = resultado.decode('utf-8').strip()
        return fecha
    except Exception as e:
        return f"No disponible: {e}"

version_codigo = obtener_fecha_ultima_actualizacion()
print("Última actualización del código:", version_codigo)

# ==========================================================

import os
import io
import threading
import time
import datetime
import tkinter as tk
import PyPDF2
import pdfplumber
import platform
import unicodedata
from PIL import Image
import pytesseract
import json
import requests
import numpy as np

import gspread
from gspread import Cell
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.service_account import Credentials
import re
from dotenv import load_dotenv
from registros import registrar_accion, registrar_error

load_dotenv()

SERVICE_ACCOUNT_FILE = 'credenciales_drive.json'
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets']

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
cliente = gspread.authorize(credentials)
service = build('drive', 'v3', credentials=credentials)

import base64
import struct
import json
def _wh_fetch_intercept_thread(debugger_address):
    """Hilo daemon que inyecta 'company-custom-bots' solo en wh.company.com."""
    import urllib.request as _ur
    import socket as _sock
    import struct as _st

    stop = threading.Event()

    def run():
        try:
            with _ur.urlopen(f"http://{debugger_address}/json", timeout=5) as r:
                pages = json.loads(r.read())
        except Exception as e:
            print(f"[WH-CDPThread] No se pudo listar tabs CDP: {e}")
            return

        ws_url = next((p["webSocketDebuggerUrl"] for p in pages
                       if "webSocketDebuggerUrl" in p), None)
        if not ws_url:
            print("[WH-CDPThread] No hay webSocketDebuggerUrl disponible.")
            return

        url_body = ws_url[len("ws://"):]
        host_port, rest = url_body.split("/", 1)
        host, port_str = host_port.rsplit(":", 1)
        path = "/" + rest

        s = _sock.socket(_sock.AF_INET, _sock.SOCK_STREAM)
        try:
            s.connect((host, int(port_str)))
        except Exception as e:
            print(f"[WH-CDPThread] Conexión TCP fallida: {e}")
            return

        ws_key = base64.b64encode(os.urandom(16)).decode()
        s.sendall((
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port_str}\r\n"
            f"Upgrade: websocket\r\n"
            f"Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {ws_key}\r\n"
            f"Sec-WebSocket-Version: 13\r\n\r\n"
        ).encode())

        buf = b""
        while b"\r\n\r\n" not in buf:
            chunk = s.recv(4096)
            if not chunk:
                break
            buf += chunk
        if b"101" not in buf:
            print("[WH-CDPThread] WebSocket handshake fallido.")
            s.close()
            return

        s.settimeout(None)
        cmd_id = [0]
        send_lock = threading.Lock()

        def ws_send(msg):
            data = json.dumps(msg).encode("utf-8")
            mask = os.urandom(4)
            n = len(data)
            if n < 126:
                header = bytes([0x81, 0x80 | n])
            elif n < 65536:
                header = bytes([0x81, 0xFE]) + _st.pack(">H", n)
            else:
                header = bytes([0x81, 0xFF]) + _st.pack(">Q", n)
            header += mask
            payload = bytes([data[i] ^ mask[i % 4] for i in range(n)])
            with send_lock:
                s.sendall(header + payload)

        def ws_recv():
            def recv_n(n):
                d = b""
                while len(d) < n:
                    c = s.recv(n - len(d))
                    if not c:
                        raise ConnectionError("CDP WebSocket cerrado.")
                    d += c
                return d
            hdr = recv_n(2)
            masked = (hdr[1] & 0x80) != 0
            length = hdr[1] & 0x7F
            if length == 126:
                length = _st.unpack(">H", recv_n(2))[0]
            elif length == 127:
                length = _st.unpack(">Q", recv_n(8))[0]
            raw = recv_n(length)
            if masked:
                mk = recv_n(4)
                raw = bytes([b ^ mk[i % 4] for i, b in enumerate(raw)])
            return json.loads(raw.decode("utf-8"))

        def next_id():
            cmd_id[0] += 1
            return cmd_id[0]

        ws_send({
            "id": next_id(),
            "method": "Fetch.enable",
            "params": {
                "patterns": [{"urlPattern": "https://wh.company.com/*",
                               "requestStage": "Request"}]
            }
        })
        print("[WH-CDPThread] Fetch intercept activo para wh.company.com.")

        while not stop.is_set():
            try:
                msg = ws_recv()
            except Exception as e:
                if not stop.is_set():
                    print(f"[WH-CDPThread] WebSocket cerrado: {e}")
                break

            if msg.get("method") == "Fetch.requestPaused":
                params = msg.get("params", {})
                request_id = params.get("requestId")
                if not request_id:
                    continue
                existing = params.get("request", {}).get("headers", {})
                headers = [{"name": k, "value": v} for k, v in existing.items()]
                headers.append({"name": "company-custom-bots",
                                "value": "TOi9mDUyKcXna0"})
                try:
                    ws_send({
                        "id": next_id(),
                        "method": "Fetch.continueRequest",
                        "params": {"requestId": request_id, "headers": headers}
                    })
                except Exception:
                    try:
                        ws_send({"id": next_id(), "method": "Fetch.continueRequest",
                                 "params": {"requestId": request_id}})
                    except Exception:
                        pass

        try:
            s.close()
        except Exception:
            pass

    t = threading.Thread(target=run, daemon=True, name="wh-cdp-fetch-intercept")
    t.start()
    return stop
_wh_intercept_stop = None
def set_company_custom_header(driver, enabled=True):
    """Activa el hilo CDP que inyecta company-custom-bots solo en wh.company.com.
    Llamar una vez tras webdriver.Chrome(). Compatible con cualquier versión de Selenium."""
    global _wh_intercept_stop
    if not enabled:
        return
    try:
        debugger_address = (driver.capabilities
                            .get("goog:chromeOptions", {})
                            .get("debuggerAddress", ""))
        if not debugger_address:
            raise ValueError("debuggerAddress no encontrado en capabilities.")
        if _wh_intercept_stop is not None:
            _wh_intercept_stop.set()
        _wh_intercept_stop = _wh_fetch_intercept_thread(debugger_address)
    except Exception as e:
        print(f"[WH-Header] CDP thread fallido ({e}). Usando fallback global (puede afectar reCAPTCHA).")
        try:
            driver.execute_cdp_cmd("Network.enable", {})
            driver.execute_cdp_cmd("Network.setExtraHTTPHeaders",
                                   {"headers": {"company-custom-bots": "TOi9mDUyKcXna0"}})
        except Exception as e2:
            print(f"[WH-Header] Fallback global tambien fallo: {e2}")
def open_wh_login_page(driver):
    """Navega a la pantalla de login de wh.company.com.
    El hilo CDP (set_company_custom_header) inyecta el header automáticamente."""
    wh_url = "https://wh.company.com/back-office"
    driver.get(wh_url)
    try:
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.ID, "email")))
        print("WH login visible.")
        return
    except Exception:
        print("WH login no visible. Reintentando via bypass...")

    driver.get("https://company.com/bypass_com_uy_on.php")
    time.sleep(1)
    driver.get(wh_url)
    try:
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.ID, "email")))
        print("WH login visible via bypass.")
    except Exception as e:
        print(f"WH login no aparecio tras segundo intento: {e}")


def login(status_label):
    global username_entry,usuario_redshift_var,usuario_redshift, contrasena_redshift,contrasena_redshift_var, usuario_db, contrasena_db, usuario_db_entry,password_db_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, username_company_entry, password_company_entry, usuario_company, contrasena_company, action_var, usuario_company, contrasena_company, username_company_entry, password_company_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, action_var, frame_middle, frame_right, frame_right_2, frame_left

    usuario = username_entry.get()
    contrasena = password_entry.get()

    usuario_company = username_company_entry.get()
    contrasena_company = password_company_entry.get()
    usuario_company = username_company_entry.get()
    contrasena_company = password_company_entry.get()

    usuario_redshift = usuario_redshift_var.get()
    contrasena_redshift = contrasena_redshift_entry.get()

    action_var.set("login_exitoso")

    thread = threading.Thread(target=update_options)
    thread.start()

def execute_code(status_label):
    global current_execution,  usuario_db,usuario_redshift_var,contrasena_redshift_var, contrasena_db, action_var, seller, username_company_entry, password_company_entry, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, log_text, valor_colum, select_option_var, valores_coma_separada, valores_coma_separada, valores_coma_separada_columna1, valores_coma_separada_columna2, valores_coma_separada_columna3

    if action_var.get() == "asignar_Courier6":
        thread = threading.Thread(target=asignar_Courier6_thread)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "extraer_pdf":
        thread = threading.Thread(target=extraer_pdf)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "extraer_nombre_pdf":
        thread = threading.Thread(target=extraer_nombre_pdf)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "facturas_Courier7":
        thread = threading.Thread(target=facturas_Courier7)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "Courier8":
        thread = threading.Thread(target=Courier8)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "facturas_Courier4":
        thread = threading.Thread(target=facturas_Courier4)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "asignar_track_MA":
        thread = threading.Thread(target=asignar_track_MA)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "cuil":
        thread = threading.Thread(target=cuil_correccion)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "facturas_oms_Courier4":
        thread = threading.Thread(target=facturas_oms_Courier4)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "company_refurbish":
        thread = threading.Thread(target=company_refurbish)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "taxes":
        thread = threading.Thread(target=taxes_thread)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")
    elif action_var.get() == "pretaxes":
        thread = threading.Thread(target=pretaxes_thread)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

    elif action_var.get() == "descripciones_cr":
        thread = threading.Thread(target=descripciones_cr_thread)
        thread.start()
        log_text.delete('1.0', tk.END)
        log_text.insert(tk.END, f"Registro de acciones:\n")

def _dcr_get_text_safe(driver, by, value, fallback=""):
    from selenium.common.exceptions import NoSuchElementException
    try:
        return driver.find_element(by, value).text.strip()
    except NoSuchElementException:
        return fallback

def _dcr_scrape_parther1(driver):
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

    product_title = _dcr_get_text_safe(driver, By.ID, "productTitle", "No se pudo obtener el título.")
    bullets = _dcr_get_text_safe(driver, By.ID, "feature-bullets", "No se pudieron obtener características.")
    details = _dcr_get_text_safe(driver, By.ID, "productDetails_techSpec_section_1", "Sin detalles adicionales.")
    details = "\n".join([
        l for l in details.split("\n")
        if not any(k in l.lower() for k in ["informar", "comentarios", "collapse", "expand"])
    ]).strip()

    tabla = []
    try:
        for fila in driver.find_elements(By.XPATH, "//tr[th[contains(@class, 'prodDetSectionEntry')]]"):
            try:
                th = fila.find_element(By.XPATH, "./th").text.strip()
                td = fila.find_element(By.XPATH, "./td").text.strip()
                if th and td:
                    tabla.append(f"{th}: {td}")
            except NoSuchElementException:
                continue
    except Exception:
        pass
    tabla_texto = "\n".join(tabla) if tabla else "Sin detalles adicionales."

    descripcion = (
        f"Título del Producto:\n{product_title}\n\n"
        f"Características:\n{bullets}\n\n"
        f"Detalles Técnicos:\n{details}\n\n"
        f"Detalles Adicionales:\n{tabla_texto}"
    )

    imagen = ""
    try:
        imagen = driver.find_element(By.ID, "landingImage").get_attribute("src") or ""
    except Exception:
        pass
    return descripcion, imagen

def _dcr_scrape_parther2(driver):
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

    product_title = ""
    for by, sel in [
        (By.CSS_SELECTOR, "h1.x-item-title__mainTitle"),
        (By.CSS_SELECTOR, ".x-item-title__mainTitle"),
        (By.ID, "itemTitle"),
    ]:
        try:
            t = driver.find_element(by, sel).text.strip()
            t = t.replace("Details about\xa0", "").replace("Details about  ", "").strip()
            if t:
                product_title = t
                break
        except NoSuchElementException:
            continue
    if not product_title:
        product_title = "No se pudo obtener el título."

    specs = []
    try:
        labels = driver.find_elements(By.CSS_SELECTOR, ".ux-labels-values__labels")
        values = driver.find_elements(By.CSS_SELECTOR, ".ux-labels-values__values")
        for lbl, val in zip(labels, values):
            l, v = lbl.text.strip(), val.text.strip()
            if l and v:
                specs.append(f"{l}: {v}")
    except Exception:
        pass
    if not specs:
        try:
            for row in driver.find_elements(By.CSS_SELECTOR, ".itemAttr tr"):
                cols = row.find_elements(By.TAG_NAME, "td")
                for i in range(0, len(cols) - 1, 2):
                    l = cols[i].text.strip().rstrip(":")
                    v = cols[i + 1].text.strip()
                    if l and v:
                        specs.append(f"{l}: {v}")
        except Exception:
            pass
    specs_texto = "\n".join(specs) if specs else "Sin especificaciones."

    desc_vendedor = ""
    try:
        iframe = driver.find_element(By.ID, "desc_ifr")
        driver.switch_to.frame(iframe)
        desc_vendedor = driver.find_element(By.TAG_NAME, "body").text[:1000].strip()
        driver.switch_to.default_content()
    except Exception:
        driver.switch_to.default_content()
        try:
            desc_vendedor = driver.find_element(By.CSS_SELECTOR, ".viTabs_content").text[:1000].strip()
        except Exception:
            pass

    descripcion = f"Título del Producto:\n{product_title}\n\nEspecificaciones:\n{specs_texto}"
    if desc_vendedor:
        descripcion += f"\n\nDescripción del Vendedor:\n{desc_vendedor}"

    imagen = ""
    for by, sel in [
        (By.CSS_SELECTOR, ".ux-image-carousel-item.active img"),
        (By.CSS_SELECTOR, ".ux-image-carousel-item img"),
        (By.ID, "icImg"),
        (By.CSS_SELECTOR, ".img.img300"),
    ]:
        try:
            src = driver.find_element(by, sel).get_attribute("src") or ""
            if src.startswith("http"):
                imagen = src
                break
        except Exception:
            continue
    return descripcion, imagen

def _dcr_scrape_producto(driver, enlace):
    import time
    driver.get(enlace)
    time.sleep(3)
    try:
        if "parther1" in enlace:
            return _dcr_scrape_parther1(driver)
        elif "parther2" in enlace:
            return _dcr_scrape_parther2(driver)
        else:
            body = driver.find_element_by_tag_name("body").text[:500].strip()
            return body, ""
    except Exception as e:
        return f"Error al obtener descripción: {e}", ""

def _dcr_crear_pdf(numero_orden, items, output_dir):
    import requests as _requests
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas as rl_canvas
    from PIL import Image as PILImage
    import io

    archivo_pdf = os.path.join(output_dir, f"Especificaciones_{numero_orden}.pdf")
    c = rl_canvas.Canvas(archivo_pdf, pagesize=letter)
    width, height = letter

    for idx, item in enumerate(items, start=1):
        if idx > 1:
            c.showPage()
        descripcion = item["descripcion"]
        enlace_imagen = item["enlace_imagen"]
        enlace = item["enlace"]
        lines_desc = descripcion.split("\n")
        product_title = lines_desc[1] if len(lines_desc) > 1 else ""

        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, f"Orden {numero_orden}")
        if len(items) > 1:
            c.setFont("Helvetica-Bold", 11)
            c.drawString(50, height - 70, f"Item {idx}/{len(items)}: {enlace[:80]}")

        y_title = height - 110 if len(items) > 1 else height - 100
        c.setFont("Helvetica-Bold", 13)
        if len(product_title) > 60:
            y_pos = y_title
            for chunk in [product_title[i:i+60] for i in range(0, len(product_title), 60)]:
                c.drawString(50, y_pos, chunk)
                y_pos -= 20
            text_y = y_pos - 20
        else:
            c.drawString(50, y_title, product_title)
            text_y = y_title - 40

        text_obj = c.beginText(50, text_y)
        text_obj.setFont("Helvetica", 10)
        max_chars = 90
        body_lines = []
        for line in "\n".join(lines_desc[2:]).split("\n"):
            while len(line) > max_chars:
                body_lines.append(line[:max_chars])
                line = line[max_chars:]
            body_lines.append(line)

        for line in body_lines:
            if any(h in line for h in ["Características", "Detalles", "Especificaciones", "Descripción"]):
                text_obj.setFont("Helvetica-Bold", 10)
            else:
                text_obj.setFont("Helvetica", 10)
            text_obj.textLine(line)
            text_y -= 14
            if text_y < 150:
                c.drawText(text_obj)
                c.showPage()
                text_y = height - 50
                text_obj = c.beginText(50, text_y)
        c.drawText(text_obj)

        if enlace_imagen and "No se pudo" not in enlace_imagen:
            try:
                resp = _requests.get(enlace_imagen, stream=True, timeout=15)
                if resp.status_code == 200:
                    img_data = PILImage.open(io.BytesIO(resp.content))
                    if img_data.mode in ("RGBA", "P"):
                        img_data = img_data.convert("RGB")
                    img_data.thumbnail((300, 300))
                    # Usar ImageReader en memoria para evitar que reportlab
                    # cachee la imagen por nombre de archivo (causa que todos
                    # los items muestren la misma imagen)
                    from reportlab.lib.utils import ImageReader
                    buf = io.BytesIO()
                    img_data.save(buf, "JPEG")
                    buf.seek(0)
                    image_reader = ImageReader(buf)
                    image_y = max(text_y - 220, 50)
                    c.drawImage(image_reader, 50, image_y, width=200, height=200)
            except Exception:
                pass
    c.save()
    return archivo_pdf

def _dcr_get_downloads_dir():
    import platform
    import pathlib
    home = pathlib.Path.home()
    system = platform.system()
    if system == "Windows":
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            ) as key:
                path = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                return pathlib.Path(path)
        except Exception:
            pass
    if system == "Darwin":
        try:
            import suomsrocess
            result = suomsrocess.check_output(
                ["osascript", "-e", "POSIX path of (path to downloads folder)"],
                text=True
            ).strip().rstrip("/")
            p = pathlib.Path(result)
            if p.exists():
                return p
        except Exception:
            pass
    # Nombres comunes según idioma del SO
    for name in ["Downloads", "Descargas", "Téléchargements", "Загрузки", "下载"]:
        candidate = home / name
        if candidate.exists():
            return candidate
    return home

def descripciones_cr_thread():
    global log_text, cliente
    import time

    def log(msg):
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        print(msg)

    output_dir = os.path.join(_dcr_get_downloads_dir(), "pdf_generados")
    os.makedirs(output_dir, exist_ok=True)

    log("Conectando a Google Sheets...")
    try:
        hoja = cliente.open("Descripciones CR").sheet1
        registros = hoja.get_all_records()
        log(f"{len(registros)} registros encontrados.")
    except Exception as e:
        log(f"Error al leer la hoja: {e}")
        return

    log("Iniciando Chrome...")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
        except Exception:
            driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(60)
    except Exception as e:
        log(f"Error al iniciar Chrome: {e}")
        return

    try:
        for registro in registros:
            numero_orden = registro.get("Numero de Orden") or registro.get("Número de Orden", "")
            enlace_campo = str(registro.get("Enlace del Producto") or "").strip()

            if not numero_orden or not enlace_campo:
                log(f"Registro incompleto, saltando: {registro}")
                continue

            enlaces = [e.strip() for e in enlace_campo.split(",") if e.strip()]
            log(f"\nOrden {numero_orden} — {len(enlaces)} item(s)...")

            items = []
            for enlace in enlaces:
                log(f"  Scrapeando: {enlace[:80]}")
                descripcion, enlace_imagen = _dcr_scrape_producto(driver, enlace)
                items.append({"enlace": enlace, "descripcion": descripcion, "enlace_imagen": enlace_imagen})

            pdf_path = _dcr_crear_pdf(numero_orden, items, output_dir)
            log(f"  PDF generado: {pdf_path}")

        log("\nProceso completado.")
    except Exception as e:
        log(f"Error inesperado: {e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

def cuil_correccion():
    global select_option_var,usuario_db_entry,password_db_entry, usuario_db, contrasena_db, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find
    import io
    import os
    import sys
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    from PIL import Image

    # Librerías portátiles
    import pypdfium2 as pdfium
    import zxingcpp

    SERVICE_ACCOUNT_FILE = 'credenciales_drive.json'
    SCOPES = ['https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets']

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"Error: Falta {SERVICE_ACCOUNT_FILE}")
        sys.exit(1)
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=credentials)

    selected_column = action_var.get()
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
    agente = agente_inconv_var.get()
    print("Agente seleccionado:", agente)

    spreadsheet_id = "1qct1Izr6lnBw0tAfYRVbEVYsTBSFlIvJO2CHwaKAKws"
    accionar = cliente.open_by_key(spreadsheet_id).worksheet("Accionar BOT")
    historial = cliente.open_by_key(spreadsheet_id).worksheet("HISTORIAL")

    valores_columna1 = accionar.col_values(1)[1:]
    valores_no_vacios_columna1 = list(dict.fromkeys([valor.strip() for valor in valores_columna1 if valor.strip()]))

    valores_columna2 = accionar.col_values(2)[1:]
    valores_no_vacios_columna2 = list(dict.fromkeys([valor.strip() for valor in valores_columna2 if valor.strip()]))

    # Verificar los valores restantes
    print(f"Valores ref después de la verificación: {valores_no_vacios_columna1}")
    print(f"Valores CUIL después de la verificación: {valores_no_vacios_columna2}")
    try:
        # Configuración de Selenium
        chrome_options = Options()
        selenium_service = None
        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(
                service=Service(driver_path),
                options=chrome_options
            )
        except Exception:
            driver = webdriver.Chrome(
                service=selenium_service,
                options=chrome_options
            )

        set_company_custom_header(driver)

        lista_valores_columna1 = valores_no_vacios_columna1
        lista_valores_columna2 = valores_no_vacios_columna2

        log_text.insert(tk.END,f"CUIL obtenidos: {lista_valores_columna2}\nRef obtenidos: {lista_valores_columna1}\n")
        print(f"CUIL obtenidos: {lista_valores_columna2}\nRef obtenidos: {lista_valores_columna1}\n")

        if not lista_valores_columna1:
            print("No hay referencias obtenidas. Deteniendo la ejecución.")
            log_text.insert(tk.END, "No hay ref nuevos. Deteniendo la ejecución.\n")
            return
        driver.get("https://company.com/bypass_com_uy_on.php")
        time.sleep(1)

        # Página web oms
        driver.get("https://use1.omsapp.com/admin_login.php?clients_id=company")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(5)

        # Página web company
        open_wh_login_page(driver)

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID,
                                            'email')))
        username.click()
        username.clear()
        username.send_keys(usuario_company)

        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID,
                                            'pass')))
        password.click()
        password.clear()
        password.send_keys(contrasena_company)
        time.sleep(1)

        button = driver.find_element(By.ID, 'send2')
        button.click()
        time.sleep(10)

        for ref, cuil in zip(lista_valores_columna1, lista_valores_columna2):
            # Limpiar CUIL proveniente de la hoja (puede venir con nombre y apellido)
            cuil_limpio = cuil.strip().split()[0] if cuil and cuil.strip() else ""
            log_text.insert(tk.END, f"Trabajando en orden:{ref}\n")
            print(f"Trabajando en orden:{ref}")

            try:
                # Primera navegación: obtener imágenes del manifest y quitar las que no están en oms
                driver.get("https://wh.company.com/edit_manifest")
                order = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, "searchBox")))
                order.clear()
                order.send_keys(ref)

                button = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, 'searchButton'))).click()
                
                time.sleep(5)

                input_dni = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "customer_dni"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_dni)

                # Intento normal: clear + send_keys
                try:
                    input_dni.clear()
                    input_dni.send_keys(cuil_limpio)
                except Exception:
                    # Fallback robusto: setear por JS y disparar eventos
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                        input_dni, cuil_limpio
                    )
                time.sleep(1)

                boton = driver.find_element(By.XPATH, "//button[@name='editar']")
                driver.execute_script("arguments[0].click();", boton)

                time.sleep(3)

            except Exception as e:
                print(f"Error al procesar la orden {ref}: {str(e)}")
                log_text.insert(tk.END, f"Error al procesar la orden {ref}\n")

                columna_accionar = accionar.col_values(3)
                fila_vacia_accionar= len(columna_accionar) + 1
                accionar.update_cell(fila_vacia_accionar, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                accionar.update_cell(fila_vacia_accionar, 4, f'{cuil_limpio}')

                columna_historial = historial.col_values(1)
                fila_vacia_historial= len(columna_historial) + 1
                fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
                agente = agente_inconv_var.get()
                historial.update_cell(fila_vacia_historial, 1, f'{ref}')
                historial.update_cell(fila_vacia_historial, 2, f'{cuil_limpio}')
                historial.update_cell(fila_vacia_historial, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                historial.update_cell(fila_vacia_historial, 4, f'{fecha_actual}')
                historial.update_cell(fila_vacia_historial, 5, f'{agente}')
                historial.update_cell(fila_vacia_historial, 7, f'{cuil}')

                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id=ref,
                    exception=e
                )

                continue
            
            try:
                driver.get("https://wh.company.com/edit_manifest")
                order = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, "searchBox")))
                order.clear()
                order.send_keys(ref)

                button = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, 'searchButton'))).click()
                
                time.sleep(5)

                input_dni = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "customer_dni"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_dni)

                valor_input_dni = input_dni.get_attribute("value")

                if valor_input_dni == cuil_limpio:
                    print(f"CUIL actualizado correctamente para la orden {ref}")
                    log_text.insert(tk.END, f"CUIL actualizado correctamente para la orden {ref}\n")

                    columna_accionar = accionar.col_values(3)
                    fila_vacia_accionar= len(columna_accionar) + 1
                    accionar.update_cell(fila_vacia_accionar, 3, f'OK - {ref} cuil: {cuil_limpio}')
                    accionar.update_cell(fila_vacia_accionar, 4, f'{cuil_limpio}')

                    columna_historial = historial.col_values(1)
                    fila_vacia_historial= len(columna_historial) + 1
                    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
                    agente = agente_inconv_var.get()
                    historial.update_cell(fila_vacia_historial, 1, f'{ref}')
                    historial.update_cell(fila_vacia_historial, 2, f'{cuil_limpio}')
                    historial.update_cell(fila_vacia_historial, 3, f'OK - {ref} cuil: {cuil_limpio}')
                    historial.update_cell(fila_vacia_historial, 4, f'{fecha_actual}')
                    historial.update_cell(fila_vacia_historial, 5, f'{agente}')
                    historial.update_cell(fila_vacia_historial, 7, f'{cuil}')
                else:
                    print(f"Error: El CUIL no se actualizó correctamente para la orden {ref}")
                    log_text.insert(tk.END, f"Error: El CUIL no se actualizó correctamente para la orden {ref}\n")

                    columna_accionar = accionar.col_values(3)
                    fila_vacia_accionar= len(columna_accionar) + 1
                    accionar.update_cell(fila_vacia_accionar, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                    accionar.update_cell(fila_vacia_accionar, 4, f'{cuil_limpio}')

                    columna_historial = historial.col_values(1)
                    fila_vacia_historial= len(columna_historial) + 1
                    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
                    agente = agente_inconv_var.get()
                    historial.update_cell(fila_vacia_historial, 1, f'{ref}')
                    historial.update_cell(fila_vacia_historial, 2, f'{cuil_limpio}')
                    historial.update_cell(fila_vacia_historial, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                    historial.update_cell(fila_vacia_historial, 4, f'{fecha_actual}')
                    historial.update_cell(fila_vacia_historial, 5, f'{agente}')
                    historial.update_cell(fila_vacia_historial, 7, f'{cuil}')

            except Exception as e:
                print(f"Error al verificar la orden {ref}: {str(e)}")
                log_text.insert(tk.END, f"Error al verificar la orden {ref}\n")

                columna_accionar = accionar.col_values(3)
                fila_vacia_accionar= len(columna_accionar) + 1
                accionar.update_cell(fila_vacia_accionar, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                accionar.update_cell(fila_vacia_accionar, 4, f'{cuil_limpio}')

                columna_historial = historial.col_values(1)
                fila_vacia_historial= len(columna_historial) + 1
                fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
                agente = agente_inconv_var.get()
                historial.update_cell(fila_vacia_historial, 1, f'{ref}')
                historial.update_cell(fila_vacia_historial, 2, f'{cuil_limpio}')
                historial.update_cell(fila_vacia_historial, 3, f'ERROR - {ref} cuil: {cuil_limpio}')
                historial.update_cell(fila_vacia_historial, 4, f'{fecha_actual}')
                historial.update_cell(fila_vacia_historial, 5, f'{agente}')
                historial.update_cell(fila_vacia_historial, 7, f'{cuil}')
                
                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id=ref,
                    exception=e
                )

                continue
        
    except Exception as e:
        log_text.insert(tk.END, f"Error en CUIL/CUIT: {e}\n")
        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )
    finally:
        log_text.insert(tk.END, "Proceso finalizado.\n")

        if driver:
            driver.quit()

        registrar_accion(
            total_orders=len(valores_no_vacios_columna1),
            agente=agente_inconv_var.get(),
        )

        # Ejecutar verificación de CUIL/CUIT al terminar
        try:
            verificar_cuil_historial()
        except Exception as e:
            if 'log_text' in globals() and log_text:
                log_text.insert(tk.END, f"Error en verificador de CUIL/CUIT: {e}\n")
            print(f"Error en verificador de CUIL/CUIT: {e}")

def verificar_cuil_historial():
    """Marca TRUE/FALSE en la columna F de HISTORIAL si ref y cuil están presentes en la columna C."""
    spreadsheet_id = "1qct1Izr6lnBw0tAfYRVbEVYsTBSFlIvJO2CHwaKAKws"
    historial = cliente.open_by_key(spreadsheet_id).worksheet("HISTORIAL")

    try:
        filas = historial.get_all_values()
        updates: list[Cell] = []

        # Saltar encabezado (fila 1) y procesar desde la fila 2
        for idx, fila in enumerate(filas[1:], start=2):
            valor_col_f = fila[5].strip().upper() if len(fila) > 5 else ""
            if valor_col_f in ("TRUE", "FALSE"):
                continue

            ref = fila[0].strip() if len(fila) > 0 else ""
            cuil = fila[1].strip() if len(fila) > 1 else ""
            texto_col_c = fila[2] if len(fila) > 2 else ""

            # Si falta información, marcamos FALSE
            if not ref or not cuil or not texto_col_c:
                updates.append(Cell(row=idx, col=6, value="FALSE"))
                continue

            coincide_ref = ref in texto_col_c
            coincide_cuil = cuil in texto_col_c

            updates.append(Cell(row=idx, col=6, value="TRUE" if (coincide_ref and coincide_cuil) else "FALSE"))

        if updates:
            historial.update_cells(updates)
            if 'log_text' in globals() and log_text:
                log_text.insert(tk.END, f"Verificador CUIL/CUIT: filas actualizadas {len(updates)}\n")
        else:
            if 'log_text' in globals() and log_text:
                log_text.insert(tk.END, "Verificador CUIL/CUIT: sin filas pendientes\n")
    except Exception as e:
        if 'log_text' in globals() and log_text:
            log_text.insert(tk.END, f"Error en verificador CUIL/CUIT: {e}\n")
        raise

def obtener_carpeta_descargas():
    home_dir = os.path.expanduser("~")
    posibles = ["Downloads", "Descargas"]
    for carpeta in posibles:
        ruta = os.path.join(home_dir, carpeta)
        if os.path.isdir(ruta):
            return ruta
    # Si no existe, crea "Downloads" por defecto
    ruta_defecto = os.path.join(home_dir, "Downloads")
    os.makedirs(ruta_defecto, exist_ok=True)
    return ruta_defecto

def obtener_servicio_drive():
    import io
    import os
    import sys
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    from PIL import Image

    # Librerías portátiles
    import pypdfium2 as pdfium
    import zxingcpp

    SERVICE_ACCOUNT_FILE = 'credenciales_drive.json'
    SCOPES = ['https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets']

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"Error: Falta {SERVICE_ACCOUNT_FILE}")
        sys.exit(1)
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=credentials)

def descargar_pdf_en_memoria(service, file_id):
    import io
    import os
    import sys
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    from PIL import Image

    # Librerías portátiles
    import pypdfium2 as pdfium
    import zxingcpp
    """
    Descarga el PDF y lo guarda en una variable en RAM (buffer).
    NO genera archivo en disco.
    """
    try:
        request = service.files().get_media(fileId=file_id)
        # Creamos un buffer en memoria en lugar de abrir un archivo
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        print("Descargando a memoria RAM...")
        while not done:
            status, done = downloader.next_chunk()
        
        # Rebobinamos el buffer al inicio para poder leerlo
        fh.seek(0)
        return fh
    except Exception as e:
        print(f"Error descarga: {e}")
        return None

def extraer_tracking_id(pdf_stream):
    import io
    import os
    import sys
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    from PIL import Image

    # Librerías portátiles
    import pypdfium2 as pdfium
    import zxingcpp
    """
    Recibe el PDF en memoria, busca el código y devuelve SOLO el string del tracking.
    """
    if not pdf_stream:
        return None

    try:
        # Cargamos el PDF directamente desde la memoria
        pdf = pdfium.PdfDocument(pdf_stream)
        page = pdf[0]
        # Renderizamos a imagen en memoria
        bitmap = page.render(scale=3)
        pil_image = bitmap.to_pil()
        
        # Leemos códigos
        results = zxingcpp.read_barcodes(pil_image)
        
        if not results:
            return None

        # --- LÓGICA DE FILTRADO ---
        # Buscamos el código que parece un tracking (Code128)
        # y evitamos el QR complejo.
        tracking_encontrado = None
        
        for result in results:
            texto = result.text.strip()
            formato = str(result.format)
            
            # Prioridad: Code128 (Barras normal)
            if "Code128" in formato:
                tracking_encontrado = texto
                break 
        
        # Si no hay Code128, devolvemos el primero que haya (fallback)
        if not tracking_encontrado and results:
            tracking_encontrado = results[0].text
            
        return tracking_encontrado

    except Exception as e:
        print(f"Error procesando imagen: {e}")
        return None
    finally:
        # Cerramos el stream para liberar memoria explícitamente (opcional pero buena práctica)
        if pdf_stream:
            pdf_stream.close()

def Courier8():
    global select_option_var,usuario_db_entry,password_db_entry, usuario_db, contrasena_db, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find
    import io
    import os
    import sys
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    from PIL import Image

    # Librerías portátiles
    import pypdfium2 as pdfium
    import zxingcpp

    # --- Configuración ---
    SCOPES = ['https://www.googleapis.com/auth/drive']

    selected_column = action_var.get()
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
    agente = agente_inconv_var.get()
    print("Agente seleccionado:", agente)

    spreadsheet_id = "1V5QCYj0x3FiTGHCmlp2wpVN_sOPKWrAKsmcNuDAy3Ac"
    changes_notas = cliente.open_by_key(spreadsheet_id).worksheet("Worksheet1")
    historial = cliente.open_by_key(spreadsheet_id).worksheet("Historial")

    # Obtener valores de la columna 6, excluyendo el encabezado
    valores_columna1 = changes_notas.col_values(6)[1:]
    valores_no_vacios_columna1 = list(dict.fromkeys([valor.strip() for valor in valores_columna1 if valor.strip()]))

    # Obtener valores de la columna 50, excluyendo el encabezado
    valores_columna2 = changes_notas.col_values(50)[1:]
    valores_no_vacios_columna2 = list(dict.fromkeys([valor.strip() for valor in valores_columna2 if valor.strip()]))

    # Obtener valores de la columna 2 historial, excluyendo el encabezado
    valores_historial = historial.col_values(2)[1:]
    valores_historial2 = list(dict.fromkeys([valor.strip() for valor in valores_historial if valor.strip()]))

    # Obtener valores de la columna 2 historial, excluyendo el encabezado
    valores_historial_so = historial.col_values(1)[1:]
    valores_historial1 = list(dict.fromkeys([valor.strip() for valor in valores_historial_so if valor.strip()]))

    DOWNLOAD_DIR = obtener_carpeta_descargas()

    # Filtrar valores de valores_no_vacios_columna1 que no están en valores_historial2
    valores_no_vacios_columna1 = [valor for valor in valores_no_vacios_columna1 if valor not in valores_historial2]

    # Filtrar valores de valores_no_vacios_columna2 que no están en valores_historial1
    valores_no_vacios_columna2 = [valor for valor in valores_no_vacios_columna2 if valor not in valores_historial1]

    # Verificar los valores restantes
    print(f"Valores ref después de la verificación: {valores_no_vacios_columna1}")
    print(f"Valores so después de la verificación: {valores_no_vacios_columna2}")

    try:
        lista_valores_columna1 = valores_no_vacios_columna1
        lista_valores_columna2 = valores_no_vacios_columna2

        log_text.insert(tk.END,
                        f"SO obtenidos: {lista_valores_columna2}\nRef obtenidos: {lista_valores_columna1}\n")
        print(f"SO obtenidos: {lista_valores_columna2}\nRef obtenidos: {lista_valores_columna1}\n")

        if not lista_valores_columna1:
            print("No hay referencias obtenidas. Deteniendo la ejecución.")
            log_text.insert(tk.END, "No hay ref nuevos. Deteniendo la ejecución.\n")
            return

        chrome_options = Options()
        chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

        # Prefs de descarga (evitan el diálogo)
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "savefile.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_settings.popups": 0
        })

        # Usar webdriver-manager para manejar ChromeDriver de forma cross-platform
        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
        except Exception:
            # Fallback a Service por defecto si webdriver-manager falla
            driver = webdriver.Chrome(service=Service(), options=chrome_options)

        set_company_custom_header(driver)

        try:
            driver.execute_cdp_cmd("Network.enable", {})
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": DOWNLOAD_DIR
            })
        except Exception:
            pass

        # Página web oms
        driver.get("https://use1.omsapp.com/admin_login.php?clients_id=company")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email_address']")))
        username.clear()
        username.send_keys(usuario)
        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password.clear()
        password.send_keys(contrasena)
        button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, 'submit-admin'))).click()

        time.sleep(5)

        # Página web mail Courier8
        driver.get("https://shipping.Courier8.com/login")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email']")))
        username.clear()
        username.send_keys(usuario_company)
        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password.clear()
        password.send_keys(contrasena_company)
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary.btn-lg.block.full-width.m-t")))
        btn.click()

        for ref,so in zip(lista_valores_columna1,lista_valores_columna2):
            log_text.insert(tk.END, f"Trabajando en oms orden:{ref} - {so}\n")
            print(f"Trabajando en oms orden:{ref} - {so}")

            file_link = []
            files_link = []

            url = f"https://shipping.Courier8.com/packages?created_from=&created_to=&tracking_code=&order_id={ref}"
            driver.get(url)
            time.sleep(5)  # Aumentar tiempo de espera
            tracking = None

            try:
                download_btn = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//button[@package_id and contains(normalize-space(.), 'Download Label')]")
                    )
                )

                package_id = download_btn.get_attribute("package_id")
                print("package_id:", package_id)
                driver.get(f"https://shipping.Courier8.com/packages/{package_id}/label")
                
                # Cerrar el cartel de descargas de Chrome
                time.sleep(2)
                try:
                    # Método 1: Presionar ESC para cerrar el cartel
                    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    print("Cartel de descargas cerrado con ESC")
                except Exception:
                    try:
                        # Método 2: JavaScript para cerrar el shelf de descargas
                        driver.execute_script("""
                            const shelf = document.querySelector('downloads-manager');
                            if (shelf && shelf.shadowRoot) {
                                const toolbar = shelf.shadowRoot.querySelector('downloads-toolbar');
                                if (toolbar && toolbar.shadowRoot) {
                                    const closeBtn = toolbar.shadowRoot.querySelector('#close');
                                    if (closeBtn) closeBtn.click();
                                }
                            }
                        """)
                        print("Cartel de descargas cerrado con JavaScript")
                    except Exception as e:
                        print(f"No se pudo cerrar el cartel de descargas: {e}")
            except Exception as e:
                print(f"No se encontró en mail Courier8, segundo intento {ref}: {str(e)}")
                log_text.insert(tk.END, f"No se encontró en mail Courier8, segundo intento {ref}\n")

                try:
                    url = f"https://shipping.Courier8.com/packages?created_from=&created_to=&tracking_code=&order_id={ref}"
                    driver.get(url)
                    time.sleep(3)

                    download_btn = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//button[@package_id and contains(normalize-space(.), 'Download Label')]")
                        )
                    )

                    package_id = download_btn.get_attribute("package_id")
                    print("package_id:", package_id)
                    driver.get(f"https://shipping.Courier8.com/packages/{package_id}/label")
                    
                    # Cerrar el cartel de descargas de Chrome
                    time.sleep(2)
                    try:
                        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                        print("Cartel de descargas cerrado con ESC")
                    except Exception:
                        try:
                            driver.execute_script("""
                                const shelf = document.querySelector('downloads-manager');
                                if (shelf && shelf.shadowRoot) {
                                    const toolbar = shelf.shadowRoot.querySelector('downloads-toolbar');
                                    if (toolbar && toolbar.shadowRoot) {
                                        const closeBtn = toolbar.shadowRoot.querySelector('#close');
                                        if (closeBtn) closeBtn.click();
                                    }
                                }
                            """)
                            print("Cartel de descargas cerrado con JavaScript")
                        except Exception as e:
                            print(f"No se pudo cerrar el cartel de descargas: {e}")

                except Exception as e:
                    print(f"No se encontró en mail Courier8 {ref}: {str(e)}")
                    log_text.insert(tk.END, f"No se encontró en mail Courier8 {ref}\n")

                    registrar_error(
                        error_message=f"Error: {str(e)}",
                        agente=agente_inconv_var.get(),
                        order_id=ref,
                        exception=e
                    )

                    columna_a_test = historial.col_values(1)
                    fila_vacia_historial = len(columna_a_test) + 1

                    historial.update_cell(fila_vacia_historial, 1, f'{so}')
                    historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                    historial.update_cell(fila_vacia_historial, 3, f'Error, no se encontro en mail Courier8')
                    historial.update_cell(fila_vacia_historial, 4, f'None')
                    historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                    historial.update_cell(fila_vacia_historial, 6, agente)
                    if tracking:
                        historial.update_cell(fila_vacia_historial, 8, tracking)
                    continue

            time.sleep(5)

            def selenium_cookies_to_requests(driver, session):
                """Copia cookies de Selenium -> requests.Session."""
                for c in driver.get_cookies():
                    # domain puede necesitar normalización; requests ignora el punto inicial
                    session.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))

            def get_pdf_url_from_perflogs(driver, timeout=10):
                """
                Lee logs de performance y devuelve la primera URL con mimeType application/pdf.
                """
                end = time.time() + timeout
                seen = set()
                while time.time() < end:
                    for entry in driver.get_log("performance"):
                        try:
                            msg = json.loads(entry["message"])["message"]
                        except Exception:
                            continue

                        if msg.get("method") == "Network.responseReceived":
                            params = msg.get("params", {})
                            resp = params.get("response", {})
                            mime = resp.get("mimeType", "")
                            url = resp.get("url")
                            if url and url not in seen:
                                seen.add(url)
                                if mime == "application/pdf":
                                    return url
                    time.sleep(0.2)
                return None

            try:
                pdf_exist = False

                OUTPUT_FILENAME = f"{ref}.pdf"
                LABEL_URL = f"https://shipping.Courier8.com/packages/{package_id}/label"

                os.makedirs(DOWNLOAD_DIR, exist_ok=True)
                output_path = os.path.join(DOWNLOAD_DIR, OUTPUT_FILENAME)

                pdf_url = get_pdf_url_from_perflogs(driver, timeout=10)

                if pdf_url:
                    s = requests.Session()
                    selenium_cookies_to_requests(driver, s)
                    headers = {
                        "User-Agent": "Mozilla/5.0",
                        "Referer": LABEL_URL
                    }
                    r = s.get(pdf_url, headers=headers, stream=True)
                    r.raise_for_status()
                    with open(output_path, "wb") as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                    print(f"PDF descargado (directo): {output_path}")

                    pdf_exist = True

                else:
                    pdf = driver.execute_cdp_cmd(
                        "Page.printToPDF",
                        {
                            "landscape": False,
                            "printBackground": True,
                            "preferCSSPageSize": True
                        }
                    )
                    import base64
                    with open(output_path, "wb") as f:
                        f.write(base64.b64decode(pdf["data"]))
                    print(f"PDF generado con printToPDF: {output_path}")

                    pdf_exist = True

            except Exception as e:
                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id=ref,
                    exception=e
                )

                print("Error al descargar PDF:", str(e))
                log_text.insert(tk.END, f"Error al descargar PDF\n")

            if pdf_exist:
                file_link = []
                files_link = []

                service = obtener_servicio_drive()

                try:
                    root_folder_id = '1TkpPijIs9ZB4DoylR4YtR6URpA1-vqGu'

                    archivo_exacto = f"{ref}.pdf"
                    etiquetas_generadas = [archivo_exacto] if archivo_exacto in os.listdir(DOWNLOAD_DIR) else []

                    if not etiquetas_generadas:
                        print(f"No se encontró el archivo {archivo_exacto} en {DOWNLOAD_DIR}")
                        log_text.insert(tk.END, f"No se encontró el archivo {archivo_exacto} en {DOWNLOAD_DIR}\n")

                        columna_a_test = historial.col_values(1)
                        fila_vacia_historial = len(columna_a_test) + 1

                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                        historial.update_cell(fila_vacia_historial, 3, f'Error, no se encontro el pdf')
                        historial.update_cell(fila_vacia_historial, 4, f'None')
                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                        historial.update_cell(fila_vacia_historial, 6, agente)
                        if tracking:
                            historial.update_cell(fila_vacia_historial, 8, tracking)
                        continue

                    etiqueta = etiquetas_generadas[0]
                    file_path = os.path.join(DOWNLOAD_DIR, etiqueta)

                    # Subir el archivo a Google Drive
                    file_metadata = {
                        'name': f"{ref}.pdf",
                        'parents': [root_folder_id]
                    }

                    media = MediaFileUpload(file_path, mimetype='application/pdf')
                    uploaded_file = service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id',
                        supportsAllDrives=True
                    ).execute()

                    file_id = uploaded_file.get('id')
                    file_link = f"https://drive.google.com/file/d/{file_id}/view"
                    files_link.append(file_link)

                    print(f'Archivo subido exitosamente. ID: {file_id}')

                    if file_link:
                        try:
                            # 1. Obtenemos el archivo en memoria (sin guardar en disco)
                            archivo_en_memoria = descargar_pdf_en_memoria(service, file_id)
                            
                            # 2. Extraemos el dato
                            tracking = extraer_tracking_id(archivo_en_memoria)
                            
                            if tracking:
                                # AQUÍ TIENES TU RESULTADO LIMPIO
                                print(f"TRACKING_ID: {tracking}")
                            else:
                                print("No se pudo extraer el tracking.")

                        except Exception as e:
                            print(f"Error al procesar y renombrar el PDF: {e}") 
                            registrar_error(
                                error_message=f"Error: {str(e)}",
                                agente=agente_inconv_var.get(),
                                order_id=ref,
                                exception=e
                            )
                            
                        try:
                            time.sleep(2)
                            # Entrar a PO y accionar
                            invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                            driver.get(invoice_url)

                            # Obtener ref
                            elemento_id = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, 'orders_customer_ref'))).get_attribute('value')
                            ref_oms = elemento_id
                            print(ref_oms)

                            if not ref == ref_oms:
                                print(f"La ref {ref} no coincide con la ref en oms {ref_oms}")
                                log_text.insert(tk.END, f"La ref {ref} no coincide con la ref en oms {ref_oms}\n")

                                columna_a_test = historial.col_values(1)
                                fila_vacia_historial = len(columna_a_test) + 1

                                historial.update_cell(fila_vacia_historial, 1, f'{so}')
                                historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                                historial.update_cell(fila_vacia_historial, 3, f'Error, ref no coincide')
                                historial.update_cell(fila_vacia_historial, 4, f'None')
                                historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                                historial.update_cell(fila_vacia_historial, 6, agente)
                                if tracking:
                                    historial.update_cell(fila_vacia_historial, 8, tracking)

                                continue

                            button_Customs = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                            button_Customs.click()

                            button_Changes = WebDriverWait(button_Customs, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Changes')]")))
                            button_Changes.click()

                            textarea = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.NAME, "PCF_CHANGES")))

                            # Obtener el contenido existente del textarea
                            existing_text = textarea.get_attribute('value')
                            textarea.clear()
                            notas_pegar = f"\n{file_link}\n"

                            # Añadir el nuevo texto al contenido existente
                            new_text = existing_text + notas_pegar
                            textarea.send_keys(new_text)

                            # Guardar
                            button_save_so = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()
                            time.sleep(2)

                            # Pegar las notas
                            button_notas = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, "//span[contains(text(), 'Notes and payment history')]")))
                            button_notas.click()

                            textarea = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.NAME, "note_text")))
                            textarea.clear()

                            resultado_notas = f"Orden viaja por Mail Courier8\nLink de la etiqueta: {file_link}\n"

                            textarea.send_keys(resultado_notas)

                            # Guardar progreso
                            button_save_clon = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                            time.sleep(2)

                            button_Custom = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                            button_Custom.click()

                            xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                            try:
                                el_shipping = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located((By.XPATH, xpath_shipping))
                                )

                                # Traer a vista
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el_shipping)
                                time.sleep(0.2)

                                # Intento normal esperando que sea clickeable
                                try:
                                    WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                    el_shipping.click()
                                except (ElementNotInteractableException, ElementClickInterceptedException,
                                        TimeoutException):
                                    # Fallback 1: click vía JS
                                    try:
                                        driver.execute_script("arguments[0].click();", el_shipping)
                                    except Exception:
                                        # Fallback 2: mover con Actions y click
                                        try:
                                            ActionChains(driver).move_to_element(el_shipping).click().perform()
                                        except Exception:
                                            # Fallback 3: focus y enviar ENTER
                                            try:
                                                driver.execute_script("arguments[0].focus();", el_shipping)
                                                el_shipping.send_keys(Keys.ENTER)
                                            except Exception as final_err:
                                                print("No se pudo hacer click en Shipping:", final_err)

                            except TimeoutException:
                                print("No se encontró el tab Shipping en la página.")

                            # Localiza el campo de entrada por su ID
                            input_field = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                            )

                            # Asegurar visibilidad antes de click
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                            time.sleep(0.3)
                            try:
                                input_field.clear()

                                input_field.send_keys(tracking)
                            except Exception:
                                # fallback por si el click normal falla
                                driver.execute_script(
                                    "arguments[0].value = arguments[1];"
                                    "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                                    "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                                    input_field, tracking
                                )

                            # Guardar progreso
                            button_save_clon = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                            try:
                                time.sleep(5)
                                # Entrar a PO y accionar
                                invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                                driver.get(invoice_url)

                                button_Custom = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located(
                                        (By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                                button_Custom.click()

                                xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                                try:
                                    el_shipping = WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located((By.XPATH, xpath_shipping))
                                    )

                                    # Traer a vista
                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el_shipping)
                                    time.sleep(0.2)

                                    # Intento normal esperando que sea clickeable
                                    try:
                                        WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                        el_shipping.click()
                                    except (ElementNotInteractableException, ElementClickInterceptedException,
                                            TimeoutException):
                                        # Fallback 1: click vía JS
                                        try:
                                            driver.execute_script("arguments[0].click();", el_shipping)
                                        except Exception:
                                            # Fallback 2: mover con Actions y click
                                            try:
                                                ActionChains(driver).move_to_element(el_shipping).click().perform()
                                            except Exception:
                                                # Fallback 3: focus y enviar ENTER
                                                try:
                                                    driver.execute_script("arguments[0].focus();", el_shipping)
                                                    el_shipping.send_keys(Keys.ENTER)
                                                except Exception as final_err:
                                                    print("No se pudo hacer click en Shipping:", final_err)

                                except TimeoutException:
                                    print("No se encontró el tab Shipping en la página.")

                                # Localiza el campo de entrada por su ID
                                input_field = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                                )

                                # Asegurar visibilidad antes de click
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                                time.sleep(0.3)

                                current_value = input_field.get_attribute('value')

                                if current_value != tracking:
                                    try:
                                        input_field.clear()

                                        input_field.send_keys(tracking)
                                    except Exception:
                                        # fallback por si el click normal falla
                                        driver.execute_script(
                                            "arguments[0].value = arguments[1];"
                                            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                                            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                                            input_field, tracking
                                        )

                                    # Guardar progreso
                                    button_save_clon = WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                                    time.sleep(5)

                                    # Entrar a PO y accionar
                                    invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                                    driver.get(invoice_url)

                                    button_Custom = WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located(
                                            (By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                                    button_Custom.click()

                                    xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                                    try:
                                        el_shipping = WebDriverWait(driver, 30).until(
                                            EC.presence_of_element_located((By.XPATH, xpath_shipping))
                                        )

                                        # Traer a vista
                                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el_shipping)
                                        time.sleep(0.2)

                                        # Intento normal esperando que sea clickeable
                                        try:
                                            WebDriverWait(driver, 10).until(
                                                EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                            el_shipping.click()
                                        except (ElementNotInteractableException, ElementClickInterceptedException,
                                                TimeoutException):
                                            # Fallback 1: click vía JS
                                            try:
                                                driver.execute_script("arguments[0].click();", el_shipping)
                                            except Exception:
                                                # Fallback 2: mover con Actions y click
                                                try:
                                                    ActionChains(driver).move_to_element(el_shipping).click().perform()
                                                except Exception:
                                                    # Fallback 3: focus y enviar ENTER
                                                    try:
                                                        driver.execute_script("arguments[0].focus();", el_shipping)
                                                        el_shipping.send_keys(Keys.ENTER)
                                                    except Exception as final_err:
                                                        print("No se pudo hacer click en Shipping:", final_err)

                                    except TimeoutException:
                                        print("No se encontró el tab Shipping en la página.")

                                    # Localiza el campo de entrada por su ID
                                    input_field = WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                                    )

                                    # Asegurar visibilidad antes de click
                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                                    time.sleep(0.3)

                                    current_value = input_field.get_attribute('value')

                                    if current_value != tracking:
                                        print(f"Orden {so} procesada con ERROR.")
                                        log_text.insert(tk.END, f"Orden {so} procesada con ERROR en track.\n")

                                        columna_a_test = historial.col_values(1)
                                        fila_vacia_historial = len(columna_a_test) + 1

                                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                                        historial.update_cell(fila_vacia_historial, 3, f'Error en track')
                                        historial.update_cell(fila_vacia_historial, 4, f'None')
                                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                                        historial.update_cell(fila_vacia_historial, 6, agente)
                                        if tracking:
                                            historial.update_cell(fila_vacia_historial, 8, tracking)

                                        continue
                                    else:
                                        print(f"Orden {so} procesada correctamente.")

                                        columna_a_test = historial.col_values(1)
                                        fila_vacia_historial = len(columna_a_test) + 1

                                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                                        historial.update_cell(fila_vacia_historial, 3, f'Ok')
                                        historial.update_cell(fila_vacia_historial, 4, f'{file_link}')
                                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                                        historial.update_cell(fila_vacia_historial, 6, agente)
                                        if tracking:
                                            historial.update_cell(fila_vacia_historial, 8, tracking)

                                else:
                                    print(f"Orden {so} procesada correctamente.")
                                    columna_a_test = historial.col_values(1)
                                    fila_vacia_historial = len(columna_a_test) + 1

                                    historial.update_cell(fila_vacia_historial, 1, f'{so}')
                                    historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                                    historial.update_cell(fila_vacia_historial, 3, f'Ok')
                                    historial.update_cell(fila_vacia_historial, 4, f'{file_link}')
                                    historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                                    historial.update_cell(fila_vacia_historial, 6, agente)
                                    if tracking:
                                        historial.update_cell(fila_vacia_historial, 8, tracking)

                            except Exception as e:
                                print(f"Error al revisar track en orden {so}", str(e))
                                log_text.insert(tk.END, f"Error al revisar track en orden {so}\n")

                                registrar_error(
                                    error_message=f"Error: {str(e)}",
                                    agente=agente_inconv_var.get(),
                                    order_id=ref,
                                    exception=e
                                )
                                columna_a_test = historial.col_values(1)
                                fila_vacia_historial = len(columna_a_test) + 1

                                historial.update_cell(fila_vacia_historial, 1, f'{so}')
                                historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                                historial.update_cell(fila_vacia_historial, 3, f'Error al revisar track en orden')
                                historial.update_cell(fila_vacia_historial, 4, f'None')
                                historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                                historial.update_cell(fila_vacia_historial, 6, agente_inconv_var.get())
                                if tracking:
                                    historial.update_cell(fila_vacia_historial, 8, tracking)

                                continue

                        except Exception as e:
                            print(f"Error al realizar orden {so}", str(e))
                            log_text.insert(tk.END, f"Error al realizar orden {so}\n")

                            registrar_error(
                                error_message=f"Error: {str(e)}",
                                agente=agente_inconv_var.get(),
                                order_id=ref,
                                exception=e
                            )
                            columna_a_test = historial.col_values(1)
                            fila_vacia_historial = len(columna_a_test) + 1

                            historial.update_cell(fila_vacia_historial, 1, f'{so}')
                            historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                            historial.update_cell(fila_vacia_historial, 3, f'Error en oms')
                            historial.update_cell(fila_vacia_historial, 4, f'None')
                            historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                            historial.update_cell(fila_vacia_historial, 6, agente_inconv_var.get())
                            if tracking:
                                historial.update_cell(fila_vacia_historial, 8, tracking)

                            continue

                    else:
                        columna_a_test = historial.col_values(1)
                        fila_vacia_historial = len(columna_a_test) + 1

                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                        historial.update_cell(fila_vacia_historial, 3, f'Error al subir el pdf')
                        historial.update_cell(fila_vacia_historial, 4, f'None')
                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                        historial.update_cell(fila_vacia_historial, 6, agente_inconv_var.get())
                        if tracking:
                            historial.update_cell(fila_vacia_historial, 8, tracking)

                        continue

                except Exception as e:
                    print(f"Error al procesar SO {so}, no se genera el pdf: {str(e)}")
                    log_text.insert(tk.END, f"Error al procesar SO {so}, no se genera el pdf: {str(e)}\n")

                    registrar_error(
                        error_message=f"Error: {str(e)}",
                        agente=agente_inconv_var.get(),
                        order_id=ref,
                        exception=e
                    )

                    columna_a_test = historial.col_values(1)
                    fila_vacia_historial = len(columna_a_test) + 1

                    historial.update_cell(fila_vacia_historial, 1, f'{so}')
                    historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                    historial.update_cell(fila_vacia_historial, 3, f'Error al generar pdf')
                    historial.update_cell(fila_vacia_historial, 4, f'None')
                    historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                    historial.update_cell(fila_vacia_historial, 6, agente)
                    if tracking:
                        historial.update_cell(fila_vacia_historial, 8, tracking)
                    continue

    except Exception as e:
        print("Error al realizar acción", str(e))
        log_text.insert(tk.END, f"Error al realizar acción\n")

        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )

    finally:
        log_text.insert(tk.END, f"Iniciando Verificación de labels\n")
        driver.quit()
        try:
            try:
                SPREADSHEET_ID = "1V5QCYj0x3FiTGHCmlp2wpVN_sOPKWrAKsmcNuDAy3Ac"
                NOMBRE_HOJA = "Historial"
                COLUMNA_B = 1
                COLUMNA_D = 3
                COLUMNA_RESULTADO = 6
                CRED_FILE = "credenciales_drive.json"

                print("🔑 Autenticando con Google...")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', "🔑 Autenticando con Google...\n")
                scopes = ["https://www.googleapis.com/auth/spreadsheets",
                          "https://www.googleapis.com/auth/drive.readonly"]
                credentials = Credentials.from_service_account_file(CRED_FILE, scopes=scopes)

                gc = gspread.authorize(credentials)
                sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(NOMBRE_HOJA)
                drive_service = build("drive", "v3", credentials=credentials)

                print("📄 Leyendo datos de la hoja...")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', "📄 Leyendo datos de la hoja...\n")
                data = sheet.get_all_values()

                cell_list = []
                print(f"📊 Total de filas en la hoja: {len(data)}")

                for i in range(1, len(data)):  # saltar encabezado
                    fila = data[i]
                    ref = fila[COLUMNA_B] if len(fila) > COLUMNA_B else ""
                    url = fila[COLUMNA_D] if len(fila) > COLUMNA_D else ""
                    resultado_actual = fila[COLUMNA_RESULTADO] if len(fila) > COLUMNA_RESULTADO else ""

                    # Solo procesar si la columna F está vacía
                    if resultado_actual:
                        continue

                    print(f"🔍 Procesando fila {i + 1}: ref={ref}, url={url}")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"🔍 Procesando fila {i + 1}: ref={ref}, url={url}\n")

                    if not url:
                        resultado = "Sin link"
                        print(f"⚠️ Sin link en fila {i + 1}")
                        if 'log_text' in globals() and log_text:
                            log_text.insert('end', f"⚠️ Sin link en fila {i + 1}\n")
                    else:
                        match = re.search(r"/d/([a-zA-Z0-9_-]{25,})", url)
                        if not match:
                            resultado = "Link inválido"
                            print(f"❌ Link inválido en fila {i + 1}")
                            if 'log_text' in globals() and log_text:
                                log_text.insert('end', f"❌ Link inválido en fila {i + 1}\n")
                        else:
                            file_id = match.group(1)
                            try:
                                file = drive_service.files().get(
                                    fileId=file_id,
                                    fields="name",
                                    supportsAllDrives=True
                                ).execute()
                                nombre = file["name"].replace(".pdf", "").strip()
                                resultado = str(nombre).lower() == str(ref).strip().lower()
                                resultado = str(resultado).upper()
                                print(f"✅ Fila {i + 1}: nombre en Drive='{nombre}', ref='{ref}', resultado={resultado}")
                                if 'log_text' in globals() and log_text:
                                    log_text.insert('end',
                                                    f"✅ Fila {i + 1}: nombre en Drive='{nombre}', ref='{ref}', resultado={resultado}\n")
                            except Exception as e:
                                if "File not found" in str(e):
                                    resultado = "Archivo no encontrado"
                                    print(f"⚠️ Archivo no encontrado en fila {i + 1}")
                                    if 'log_text' in globals() and log_text:
                                        log_text.insert('end', f"⚠️ Archivo no encontrado en fila {i + 1}\n")
                                else:
                                    resultado = f"Error: {str(e)}"
                                    print(f"⚠️ Error en fila {i + 1}: {str(e)}")
                                    if 'log_text' in globals() and log_text:
                                        log_text.insert('end', f"⚠️ Error en fila {i + 1}: {str(e)}\n")

                    # Solo agregar si la fila tiene algún dato relevante (por ejemplo, ref o url)
                    if ref or url:
                        cell = Cell(row=i + 1, col=COLUMNA_RESULTADO + 1, value=resultado)
                        cell_list.append(cell)

                if cell_list:
                    print(f"⬆️ Actualizando hoja en {len(cell_list)} celdas de la columna F...")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"⬆️ Actualizando hoja en {len(cell_list)} celdas de la columna F...\n")
                    sheet.update_cells(cell_list)

                print(f"✅ {len(cell_list)} filas procesadas.")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', f"✅ {len(cell_list)} filas procesadas.\n")

            except Exception as e:
                print("Error al realizar verificación", str(e))
                log_text.insert(tk.END, f"Error al realizar verificación\n")

                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id="Error al realizar verificacion",
                    exception=e
                )
            try:
                asignar_track_MA()
            except Exception as e:
                print("Error al realizar asignaciones faltantes", str(e))
                log_text.insert(tk.END, f"Error al realizar asignaciones faltantes\n")

                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id="Error al realizar asignaciones faltantes",
                    exception=e
                )
        except Exception as e:
            print("Error al realizar finally", str(e))
            log_text.insert(tk.END, f"Error al realizar finally\n")

            registrar_error(
                error_message=f"Error: {str(e)}",
                agente=agente_inconv_var.get(),
                order_id="Error finally",
                exception=e
            )
        finally:
            registrar_accion(
                total_orders=len(valores_no_vacios_columna1),
                agente=agente_inconv_var.get(),
            )

            # Mostrar mensaje de éxito
            log_text.insert(tk.END, f"Finalizado\n")
            driver.quit()

def get_desktop_path() -> Optional[str]:
    home = os.path.expanduser("~")
    candidates = []

    # Rutas típicas
    candidates += [
        os.path.join(home, "Desktop"),
        os.path.join(home, "Escritorio"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive", "Escritorio"),
    ]

    # Si existe variable de entorno de OneDrive, probar ahí también
    onedrive_env = os.environ.get("OneDrive") or os.environ.get("OneDriveConsumer")
    if onedrive_env:
        candidates += [
            os.path.join(onedrive_env, "Desktop"),
            os.path.join(onedrive_env, "Escritorio"),
        ]

    for p in candidates:
        if p and os.path.isdir(p):
            return p

    # Último recurso: API de Windows (funciona en Win7+)
    try:
        import ctypes, ctypes.wintypes as wt
        CSIDL_DESKTOPDIRECTORY = 0x10
        buf = ctypes.create_unicode_buffer(wt.MAX_PATH)
        if ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf) == 0:
            p = buf.value
            if os.path.isdir(p):
                return p
    except Exception:
        pass

    return None

def company_refurbish():
    global select_option_var, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find
    try:
        log_text.insert(tk.END, "Iniciando proceso company Refurbish...\n")
        # Abre la hoja de Google Sheets
        sheet = cliente.open("RPA | Control Prod Refurbish").sheet1
        # Lee todas las filas
        rows = sheet.get_all_records()
        # Inicializa las listas para guardar los valores de "order" que necesitan revisión
        revisar = []
        revisar2 = []
        # Actualiza varias celdas en batch para reducir llamadas a la API
        cell_updates = []
        for index, row in enumerate(rows, start=2):  # Comienza en 2 para compensar la fila de encabezado
            name = str(row['Name']).lower()
            if 'refurbish' in name or 'renew' in name:
                order_value = row['Order']
                revisar.append(order_value)
                so_value = row['SO']
                result2 = f"FAC.{so_value}/G.{order_value}"
                revisar2.append(result2)
                cell_updates.append({
                    'range': f'F{index}',
                    'values': [[result2]]
                })
        # Ejecuta la actualización en batch
        if cell_updates:
            sheet.batch_update(cell_updates)
        log_text.insert(tk.END, f"Order: {revisar}\n")
        log_text.insert(tk.END, f"Código: {revisar2}\n")
        # Inicializa el driver de Selenium
        chrome_options = Options()
        driverTM = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        driverTM.maximize_window()
        # Navega a la página de inicio de sesión
        # Página web company
        open_wh_login_page(driver)

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID,
                                            'email')))
        username.click()
        username.clear()
        username.send_keys(usuario_company)

        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID,
                                            'pass')))
        password.click()
        password.clear()
        password.send_keys(contrasena_company)
        time.sleep(1)

        button = driver.find_element(By.ID, 'send2')
        button.click()
        time.sleep(10)
        # Navega a la página de seguimiento
        driverTM.get("https://company.com.uy/logistictools/index/trakings")
        # Completa los campos de seguimiento para "Order"
        try:
            asignartracking_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#logistic-form > div:nth-child(2) > h5")))
            asignartracking_button.click()
        except:
            pass
        for order in revisar:
            asignartrackbox_button = driverTM.find_element(By.CSS_SELECTOR, "#logistic-form > div:nth-child(2) > div > textarea")
            asignartrackbox_button.send_keys(order)
            asignartrackbox_button.send_keys(Keys.RETURN)
            time.sleep(1)
        # Completa los campos de seguimiento para "Código"
        try:
            numtracking_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#logistic-form > div:nth-child(3) > h5")))
            numtracking_button.click()
        except:
            pass
        for codigo in revisar2:
            numtrackbox_button = driverTM.find_element(By.CSS_SELECTOR, "#logistic-form > div:nth-child(3) > div > textarea")
            numtrackbox_button.send_keys(codigo)
            numtrackbox_button.send_keys(Keys.RETURN)
            time.sleep(1)
        # Botón para enviar el tracking
        astrack_butt = driverTM.find_element(By.CSS_SELECTOR, "#logistic-form > div.form-control.col-12 > button")
        astrack_butt.click()
        log_text.insert(tk.END, f'Se agregó el motivo de retención en el # de tracking a las órdenes: {revisar}\n')
        # Accede a la herramienta de envío masivo
        driverTM.get("https://company.com.uy/logistictools/index/enviomailmasivo")
        def procesar_template(revisar):
            try:
                opa_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#logistic-form > div:nth-child(1) > h5")))
                opa_button.click()
            except:
                pass
            for order in revisar:
                opa_box = driverTM.find_element(By.CSS_SELECTOR, "#logistic-form > div:nth-child(1) > div > textarea")
                opa_box.send_keys(order)
                opa_box.send_keys(Keys.RETURN)
                time.sleep(1)
            template_box = driverTM.find_element(By.CSS_SELECTOR, "#template")
            template_box.click()
            titulo_de_correo = driverTM.find_element(By.CSS_SELECTOR, "#mailTitle")
            titulo_de_correo.send_keys('Notificación company')
            template_desp = driverTM.find_element(By.CSS_SELECTOR, "#ddtemplates")
            template_desp.click()
            template_text = "CR 4.1 - Renovados"
            select_template = driverTM.find_element(By.XPATH, f"//select[@id='ddtemplates']/option[contains(text(), '{template_text}')]")
            select_template.click()
            textarea_button = driverTM.find_element(By.CSS_SELECTOR, "#logistic-form > div:nth-child(1) > div > textarea")
            textarea_button.click()
            time.sleep(1)
            enviar_button = driverTM.find_element(By.CSS_SELECTOR, "#fcontrols > div.form-control.col-12 > button")
            enviar_button.click()
        time.sleep(4)
        procesar_template(revisar)
        log_text.insert(tk.END, 'Listo.\n')
    except Exception as e:
        log_text.insert(tk.END, f"Error en company_refurbish: {e}\n")

def facturas_Courier4():
    global select_option_var, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find

    try:
        log_text.insert(tk.END, "Proceso iniciado.\n")
        
        direccion_clave = "8404 NW 90TH ST UNIT 300TH, MIAMI, FL 33166, United States"
        desktop_folder = get_desktop_path()
        if not desktop_folder:
            log_text.insert(tk.END, "No se encontró la carpeta de escritorio.\n")
            return

        print("Escritorio encontrado en:", desktop_folder)

        archivo_pdf_gigante = os.path.join(desktop_folder, "ORDENES.pdf")
        if not os.path.isfile(archivo_pdf_gigante):
            log_text.insert(tk.END, "No se encontró `ORDENES.pdf` en el escritorio.\n")
            return

        carpeta_destino = os.path.join(desktop_folder, "Labels ECU Courier4")
        os.makedirs(carpeta_destino, exist_ok=True)

        def extraer_numero_orden(pdf_path):
            try:
                # Verificar que el archivo existe
                if not os.path.exists(pdf_path):
                    log_text.insert(tk.END, f"Archivo no existe: {pdf_path}\n")
                    return None
                
                # Intentar abrir y extraer texto
                with pdfplumber.open(pdf_path) as pdf:
                    if not pdf.pages:
                        log_text.insert(tk.END, f"PDF vacío: {pdf_path}\n")
                        return None
                    primera_pagina = pdf.pages[0].extract_text()
                
                # Verificar que se extrajo texto
                if not primera_pagina:
                    log_text.insert(tk.END, f"No se pudo extraer texto del PDF: {pdf_path}\n")
                    return None
                    
                if primera_pagina:
                    clave_busqueda = "Orden: "
                    inicio_orden = primera_pagina.find(clave_busqueda)
                    if inicio_orden != -1:
                        inicio_orden += len(clave_busqueda)
                        fin_orden = primera_pagina.find(" ", inicio_orden)
                        orden = primera_pagina[inicio_orden:fin_orden] if fin_orden != -1 else primera_pagina[inicio_orden:]
                        # Limpiar caracteres extraños
                        orden_limpia = orden.strip()
                        if orden_limpia:
                            return orden_limpia
                log_text.insert(tk.END, f"No se encontró 'Orden:' en: {pdf_path}\n")
            except PermissionError as e:
                log_text.insert(tk.END, f"Error de permisos al acceder: {pdf_path} - {e}\n")
            except Exception as e:
                log_text.insert(tk.END, f"Error al extraer número de orden de {pdf_path}: {type(e).__name__} - {e}\n")
            return None

        try:
            with open(archivo_pdf_gigante, 'rb') as f:
                lector_pdf = PyPDF2.PdfReader(f)
                total_paginas = len(lector_pdf.pages)
                if total_paginas == 0:
                    log_text.insert(tk.END, "El PDF está vacío.\n")
                    return

                escritor_pdf = PyPDF2.PdfWriter()
                paginas_acumuladas = []
                fragmento_num = 1

                for i in range(total_paginas):
                    try:
                        with pdfplumber.open(archivo_pdf_gigante) as pdf:
                            pagina_texto = pdf.pages[i].extract_text() if i < len(pdf.pages) else ""
                    except Exception as e:
                        log_text.insert(tk.END, f"Error leyendo página {i}: {e}\n")
                        continue

                    try:
                        escritor_pdf.add_page(lector_pdf.pages[i])
                        paginas_acumuladas.append(lector_pdf.pages[i])
                    except Exception as e:
                        log_text.insert(tk.END, f"Error agregando página {i}: {e}\n")
                        continue

                    if direccion_clave in pagina_texto:
                        temp_pdf_path = os.path.join(carpeta_destino, f'fragmento_{fragmento_num}.pdf')
                        try:
                            with open(temp_pdf_path, 'wb') as output:
                                escritor_pdf.write(output)
                        except Exception as e:
                            log_text.insert(tk.END, f"Error guardando fragmento {fragmento_num}: {e}\n")
                            continue

                        numero_orden = extraer_numero_orden(temp_pdf_path)
                        if numero_orden:
                            nuevo_pdf_path = os.path.join(carpeta_destino, f'{numero_orden}.pdf')
                            try:
                                os.rename(temp_pdf_path, nuevo_pdf_path)
                            except Exception as e:
                                log_text.insert(tk.END, f"Error renombrando archivo: {e}\n")
                        else:
                            log_text.insert(tk.END, f"No se pudo extraer el número de orden del archivo {temp_pdf_path}.\n")

                        escritor_pdf = PyPDF2.PdfWriter()
                        paginas_acumuladas = []
                        fragmento_num += 1

                # Guardar el último fragmento si no se ha guardado aún
                if paginas_acumuladas:
                    temp_pdf_path = os.path.join(carpeta_destino, f'fragmento_{fragmento_num}.pdf')
                    try:
                        with open(temp_pdf_path, 'wb') as output:
                            escritor_pdf.write(output)
                    except Exception as e:
                        log_text.insert(tk.END, f"Error guardando último fragmento: {e}\n")

                    numero_orden = extraer_numero_orden(temp_pdf_path)
                    if numero_orden:
                        nuevo_pdf_path = os.path.join(carpeta_destino, f'{numero_orden}.pdf')
                        try:
                            os.rename(temp_pdf_path, nuevo_pdf_path)
                        except Exception as e:
                            log_text.insert(tk.END, f"Error renombrando archivo: {e}\n")
                    else:
                        log_text.insert(tk.END, f"No se pudo extraer el número de orden del archivo {temp_pdf_path}.\n")
        except Exception as e:
            log_text.insert(tk.END, f"Error procesando el PDF gigante: {e}\n")
            registrar_error(
                error_message=f"Error: {str(e)}",
                agente=agente_inconv_var.get(),
                order_id=numero_orden,
                exception=e
            )

    except Exception as e:
        log_text.insert(tk.END, f"Error en facturas_Courier4: {e}\n")
        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )
    finally:
        log_text.insert(tk.END, "Proceso finalizado.\n")

        registrar_accion(
            total_orders=total_paginas,
            agente=agente_inconv_var.get(),
        )

def facturas_Courier7():
    global select_option_var, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find
    ordenes = []
    try:
        log_text.insert(tk.END, "Proceso iniciado.\n")
        desktop_folder = get_desktop_path()
        if not desktop_folder:
            log_text.insert(tk.END, "No se encontró la carpeta de escritorio.\n")
            return

        print("Escritorio encontrado en:", desktop_folder)

        archivo_pdf_gigante = os.path.join(desktop_folder, "ORDENES.pdf")
        if not os.path.isfile(archivo_pdf_gigante):
            log_text.insert(tk.END, "No se encontró `ORDENES.pdf` en el escritorio.\n")
            return

        carpeta_destino = os.path.join(desktop_folder, "Labels ECU Courier7")
        os.makedirs(carpeta_destino, exist_ok=True)

        def extraer_numero_orden(pdf_path):
            try:
                # Verificar que el archivo existe
                if not os.path.exists(pdf_path):
                    log_text.insert(tk.END, f"Archivo no existe: {pdf_path}\n")
                    return None
                
                # Intentar abrir y extraer texto
                with pdfplumber.open(pdf_path) as pdf:
                    if not pdf.pages:
                        log_text.insert(tk.END, f"PDF vacío: {pdf_path}\n")
                        return None
                    primera_pagina = pdf.pages[0].extract_text()
                
                # Verificar que se extrajo texto
                if not primera_pagina:
                    log_text.insert(tk.END, f"No se pudo extraer texto del PDF: {pdf_path}\n")
                    return None
                    
                if primera_pagina:
                    clave_busqueda = "Orden: "
                    inicio_orden = primera_pagina.find(clave_busqueda)
                    if inicio_orden != -1:
                        inicio_orden += len(clave_busqueda)
                        fin_orden = primera_pagina.find(" ", inicio_orden)
                        orden = primera_pagina[inicio_orden:fin_orden] if fin_orden != -1 else primera_pagina[inicio_orden:]
                        # Limpiar caracteres extraños
                        orden_limpia = orden.strip()
                        if orden_limpia:
                            return orden_limpia
                log_text.insert(tk.END, f"No se encontró 'Orden:' en: {pdf_path}\n")
            except PermissionError as e:
                log_text.insert(tk.END, f"Error de permisos al acceder: {pdf_path} - {e}\n")
            except Exception as e:
                log_text.insert(tk.END, f"Error al extraer número de orden de {pdf_path}: {type(e).__name__} - {e}\n")
            return None

        try:
            with open(archivo_pdf_gigante, 'rb') as f:
                lector_pdf = PyPDF2.PdfReader(f)
                total_paginas = len(lector_pdf.pages)
                if total_paginas == 0:
                    log_text.insert(tk.END, "El PDF está vacío.\n")
                    return

                for i in range(0, total_paginas, 2):
                    escritor_pdf = PyPDF2.PdfWriter()
                    try:
                        escritor_pdf.add_page(lector_pdf.pages[i])
                        if i + 1 < total_paginas:
                            escritor_pdf.add_page(lector_pdf.pages[i + 1])
                    except Exception as e:
                        log_text.insert(tk.END, f"Error agregando páginas en fragmento {i // 2 + 1}: {e}\n")
                        continue

                    temp_pdf_path = os.path.join(carpeta_destino, f'part_{i // 2 + 1}.pdf')
                    try:
                        with open(temp_pdf_path, 'wb') as output:
                            escritor_pdf.write(output)
                    except Exception as e:
                        log_text.insert(tk.END, f"Error guardando fragmento {i // 2 + 1}: {e}\n")
                        continue

                    numero_orden = extraer_numero_orden(temp_pdf_path)
                    if numero_orden:
                        # Verifica si termina con " S.O." y lo elimina
                        if numero_orden.endswith(" S.O."):
                            numero_orden = numero_orden[:-5].rstrip()
                        nuevo_pdf_path = os.path.join(carpeta_destino, f'{numero_orden}.pdf')
                        try:
                            os.rename(temp_pdf_path, nuevo_pdf_path)
                        except Exception as e:
                            log_text.insert(tk.END, f"Error renombrando archivo: {e}\n")
                    else:
                        log_text.insert(tk.END, f"No se pudo extraer el número de orden del archivo {temp_pdf_path}.\n")
        except Exception as e:
            log_text.insert(tk.END, f"Error procesando el PDF gigante: {e}\n")
            registrar_error(
                error_message=f"Error: {str(e)}",
                agente=agente_inconv_var.get(),
                order_id=numero_orden,
                exception=e
            )


    except Exception as e:
        log_text.insert(tk.END, f"Error en facturas_Courier7: {e}\n")
        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )

    finally:
        log_text.insert(tk.END, "Proceso finalizado.\n")
        registrar_accion(
            total_orders=total_paginas,
            agente=agente_inconv_var.get(),
        )

def recortar_bordes_blancos(img, umbral=250):
    """Recorta los bordes blancos de una imagen PIL.
    
    Args:
        img: Imagen PIL
        umbral: Valor de gris por encima del cual se considera blanco (0-255)
    
    Returns:
        Imagen PIL recortada
    """
    # Convertir a escala de grises para detectar bordes blancos
    img_gray = img.convert('L')
    # Convertir a array numpy para procesamiento
    import numpy as np
    img_array = np.array(img_gray)
    
    # Encontrar filas y columnas que no son completamente blancas
    filas_no_blancas = np.where(np.min(img_array, axis=1) < umbral)[0]
    columnas_no_blancas = np.where(np.min(img_array, axis=0) < umbral)[0]
    
    if len(filas_no_blancas) == 0 or len(columnas_no_blancas) == 0:
        # Si la imagen es completamente blanca, devolver original
        return img
    
    # Obtener límites del contenido
    top = filas_no_blancas[0]
    bottom = filas_no_blancas[-1]
    left = columnas_no_blancas[0]
    right = columnas_no_blancas[-1]
    
    # Recortar la imagen
    return img.crop((left, top, right + 1, bottom + 1))

def facturas_oms_Courier4():
    global select_option_var, usuario_db_entry, password_db_entry, usuario_db, contrasena_db, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry

    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
    agente = agente_inconv_var.get()
    print("Agente seleccionado:", agente)

    # ID de la carpeta de Drive donde se subirán los PDFs
    ROOT_FOLDER_ID = "1TkpPijIs9ZB4DoylR4YtR6URpA1-vqGu"  # Ajusta este ID según tu carpeta

    spreadsheet_id = "1NXUH5gy14uP2iKnu_mwxK_QOSnl-j-y8fcOgl00LU0I"  # Ajusta según tu sheet
    sheet = cliente.open_by_key(spreadsheet_id).worksheet("Worksheet1")
    historial = cliente.open_by_key(spreadsheet_id).worksheet("Historial")

    # Obtener todos los datos de la hoja
    all_data = sheet.get_all_values()
    

    # Procesar desde la fila 2 (índice 1) para saltar encabezado
    filas_a_procesar = []
    for idx, fila in enumerate(all_data[1:], start=2):
        so = fila[0].strip() if len(fila) > 0 else ""
        link_existente = fila[1].strip() if len(fila) > 1 else ""
        courier = fila[2].strip() if len(fila) > 2 else ""
        # Solo procesar si tiene SO y no tiene link
        if so and not link_existente:
            filas_a_procesar.append((idx, so, courier))

    print(f"SO a procesar: {[so for _, so, _ in filas_a_procesar]}")
    log_text.insert(tk.END, f"SO a procesar: {len(filas_a_procesar)} órdenes\n")

    if not filas_a_procesar:
        print("No hay SO nuevos para procesar.")
        log_text.insert(tk.END, "No hay SO nuevos para procesar.\n")
        return

    DOWNLOAD_DIR = obtener_carpeta_descargas()

    try:
        # Configuración de Selenium
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "savefile.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_settings.popups": 0
        })

        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(
                service=Service(driver_path),
                options=chrome_options
            )
        except Exception:
            driver = webdriver.Chrome(
                service=Service(), options=chrome_options)
            
        set_company_custom_header(driver)

        try:
            driver.execute_cdp_cmd("Network.enable", {})
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": DOWNLOAD_DIR
            })
        except Exception:
            pass

        # Login a oms
        driver.get("https://use1.omsapp.com/admin_login.php?clients_id=company")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        username_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email_address']")))
        username_field.clear()
        username_field.send_keys(usuario)
        
        password_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password_field.clear()
        password_field.send_keys(contrasena)
        
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, 'submit-admin'))).click()

        time.sleep(5)

        try:
            # Página web company
            driver.get("https://wh.company.com/back-office")

            # Inicio de sesión
            username = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='login[username]']")))
            username.clear()
            username.send_keys(usuario_company)
            time.sleep(1)
            password = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='login[password]']")))
            password.clear()
            password.send_keys(contrasena_company)
            time.sleep(1)
            button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, 'send2'))).click()

            time.sleep(5)

        except Exception as e:
            print(e)

        # Función helper para agregar al historial (ahora incluye courier)
        def agregar_a_historial(so, link, ref, courier):
            try:
                # Obtener la última fila con datos en Historial
                all_historial = historial.get_all_values()
                ultima_fila = len(all_historial) + 1
                # Agregar nueva fila con SO, link, ref y courier (valores primero, rango después)
                historial.update(values=[[so, link, ref, courier]], range_name=f'A{ultima_fila}:D{ultima_fila}')
                print(f"Agregado al historial: SO={so}, link={link}, ref={ref}, courier={courier}")
            except Exception as e:
                print(f"Error agregando al historial: {e}")
                log_text.insert(tk.END, f"Error agregando al historial: {e}\n")

        # Procesar cada SO
        for fila_num, so, courier in filas_a_procesar:
            log_text.insert(tk.END, f"Procesando SO: {so} (fila {fila_num})\n")
            print(f"Procesando SO: {so} (fila {fila_num})")

            try:
                # Ir a la página de Notes and Payments
                invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                driver.get(invoice_url)
                time.sleep(3)

                # Obtener ref para validación
                try:
                    elemento_ref = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'orders_customer_ref')))
                    ref_oms = elemento_ref.get_attribute('value')
                    print(f"Ref encontrada: {ref_oms}")
                except Exception as e:
                    print(f"No se pudo obtener la ref: {e}")
                    ref_oms = so

                # Click en Notes and Payments
                try:
                    notes_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Notes and payment history')]")))
                    notes_button.click()
                    time.sleep(2)
                except Exception as e:
                    print(f"No se pudo hacer click en Notes and Payments: {e}")
                    log_text.insert(tk.END, f"Error al acceder a Notes and Payments para SO {so}\n")
                    
                    error_msg = f'Error: No se pudo acceder a Notes and Payments'
                    # Marcar error en la columna B
                    sheet.update_cell(fila_num, 2, error_msg)
                    # Agregar al historial con error
                    agregar_a_historial(so, error_msg, ref_oms, courier)
                    continue

                # Buscar el enlace de factura en Notes
                factura_url = None
                try:
                    # Buscar todos los enlaces que contienen "labelv2"
                    enlaces = driver.find_elements(By.XPATH, "//a[contains(@href, 'labelv2')]")
                    
                    if enlaces:
                        for enlace in enlaces:
                            href = enlace.get_attribute('href')
                            if 'company.com/labelv2' in href:
                                factura_url = href
                                print(f"URL de factura encontrada: {factura_url}")
                                log_text.insert(tk.END, f"URL encontrada: {factura_url}\n")
                                break
                    
                    if not factura_url:
                        print(f"No se encontró enlace de factura para SO {so}")
                        log_text.insert(tk.END, f"No se encontró enlace de factura para SO {so}\n")
                        
                        error_msg = f'Error: No se encontró enlace de factura'
                        # Marcar error en la columna B
                        sheet.update_cell(fila_num, 2, error_msg)
                        # Agregar al historial con error
                        agregar_a_historial(so, error_msg, ref_oms, courier)
                        continue

                except Exception as e:
                    print(f"Error buscando enlace de factura: {e}")
                    log_text.insert(tk.END, f"Error buscando enlace: {e}\n")
                    
                    error_msg = f'Error: {str(e)}'
                    sheet.update_cell(fila_num, 2, error_msg)
                    # Agregar al historial con error
                    agregar_a_historial(so, error_msg, ref_oms, courier)
                    continue

                # Descargar el PDF usando Chrome DevTools Protocol (CDP)
                try:
                    custom_header = {"company-custom-bots": "TOi9mDUyKcXna0"}

                    # Abrir una pestaña en blanco, aplicar header y luego navegar.
                    # Esto evita captcha en dominios donde la URL de factura requiere el mismo header de WH.
                    driver.execute_script("window.open('about:blank', '_blank');")
                    time.sleep(2)
                    
                    # Cambiar a la nueva pestaña
                    driver.switch_to.window(driver.window_handles[-1])

                    try:
                        driver.execute_cdp_cmd("Network.enable", {})
                        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {
                            "headers": custom_header
                        })
                    except Exception as e:
                        print(f"No se pudo configurar header custom en pestaña de factura: {e}")

                    driver.get(factura_url)
                    time.sleep(3)  # Esperar a que cargue completamente
                    
                    # Usar Chrome DevTools Protocol para imprimir a PDF
                    print("Generando PDF usando CDP...")
                    log_text.insert(tk.END, f"Generando PDF para SO {so}...\n")
                    
                    # Configuración de impresión a PDF en tamaño 6x4 pCourier5adas
                    # 6 pCourier5adas = 15.24 cm, 4 pCourier5adas = 10.16 cm
                    pdf_settings = {
                        'landscape': False,
                        'displayHeaderFooter': False,
                        'printBackground': True,
                        'preferCSSPageSize': False,  # No usar tamaño CSS, usar el que especificamos
                        'paperWidth': 6,  # 6 pCourier5adas de ancho
                        'paperHeight': 4,  # 4 pCourier5adas de alto
                        'marginTop': 0,
                        'marginBottom': 0,
                        'marginLeft': 0,
                        'marginRight': 0,
                    }
                    
                    # Ejecutar comando de impresión a PDF
                    result = driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)
                    pdf_data = result['data']
                    
                    # Decodificar el PDF de base64
                    import base64
                    import fitz  # PyMuPDF
                    from PyPDF2 import PdfReader, PdfWriter
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.utils import ImageReader
                    from PIL import Image as PILImage
                    pdf_bytes = base64.b64decode(pdf_data)
                    
                    # Guardar temporalmente el PDF completo
                    pdf_filename_temp = f"Temp_Factura_{ref_oms}_{so}.pdf"
                    file_path_temp = os.path.join(DOWNLOAD_DIR, pdf_filename_temp)
                    
                    with open(file_path_temp, 'wb') as f:
                        f.write(pdf_bytes)
                    
                    # Extraer desde la página 3 en adelante (omitir páginas 1 y 2) y unificar en una sola página
                    pdf_filename = f"Factura_{ref_oms}_{so}.pdf"
                    file_path = os.path.join(DOWNLOAD_DIR, pdf_filename)
                    
                    try:
                        # Abrir el PDF con PyMuPDF
                        doc = fitz.open(file_path_temp)
                        total_paginas = len(doc)
                        
                        # Ancho fijo de 6 pCourier5adas
                        pdf_width = 6 * 72  # 432 puntos
                        
                        # Determinar qué páginas procesar
                        if total_paginas > 2:
                            # Convertir páginas 3+ a imágenes y unificarlas
                            print(f"Convirtiendo páginas 3-{total_paginas} a una sola hoja de 6\" de ancho...")
                            log_text.insert(tk.END, f"Unificando páginas 3-{total_paginas} en una sola\n")
                            
                            # Lista para almacenar las imágenes
                            paginas_imagenes = []
                            
                            # Convertir páginas 3+ a imágenes (índices 2+)
                            for page_num in range(2, total_paginas):
                                page = doc[page_num]
                                # Renderizar página a imagen con alta resolución para impresión
                                mat = fitz.Matrix(4, 4)  # Mayor resolución para mejor calidad de impresión
                                pix = page.get_pixmap(matrix=mat)
                                # Convertir a PIL Image
                                img_data = pix.tobytes("png")
                                img = PILImage.open(io.BytesIO(img_data))
                                # Recortar bordes blancos
                                img = recortar_bordes_blancos(img)
                                paginas_imagenes.append(img)
                            
                            doc.close()
                            
                            if paginas_imagenes:
                                # Ancho de referencia (usar el ancho del PDF en píxeles a 300 DPI para impresión)
                                # 6 pCourier5adas * 300 DPI = 1800 píxeles
                                target_width = int(pdf_width / 72 * 300)  # Convertir puntos a píxeles a 300 DPI
                                
                                # Redimensionar todas las imágenes al mismo ancho manteniendo proporción
                                imagenes_redimensionadas = []
                                for img in paginas_imagenes:
                                    # Calcular nueva altura manteniendo proporción
                                    aspect_ratio = img.height / float(img.width)
                                    new_height = int(target_width * aspect_ratio)
                                    # Redimensionar imagen
                                    img_resized = img.resize((target_width, new_height), PILImage.Resampling.LANCZOS)
                                    imagenes_redimensionadas.append(img_resized)
                                
                                # Combinar todas las imágenes redimensionadas verticalmente
                                total_height = sum(img.height for img in imagenes_redimensionadas)
                                
                                # Crear imagen combinada con el ancho objetivo
                                combined_image = PILImage.new('RGB', (target_width, total_height), 'white')
                                
                                # Pegar cada imagen una debajo de la otra
                                current_y = 0
                                for img in imagenes_redimensionadas:
                                    combined_image.paste(img, (0, current_y))
                                    current_y += img.height
                                
                                # Guardar la imagen combinada temporalmente
                                temp_image_path = os.path.join(DOWNLOAD_DIR, f"temp_combined_{so}.png")
                                combined_image.save(temp_image_path, 'PNG')
                                
                                # Calcular altura proporcional para mantener 6" de ancho
                                img_aspect = total_height / float(target_width)
                                pdf_height = pdf_width * img_aspect  # Altura proporcional
                                
                                # Crear PDF con 6" de ancho y altura variable
                                c = canvas.Canvas(file_path, pagesize=(pdf_width, pdf_height))
                                
                                # Dibujar imagen ocupando toda la página sin espacios
                                c.drawImage(temp_image_path, 0, 0, 
                                           width=pdf_width, height=pdf_height,
                                           preserveAspectRatio=False)  # Forzar a ocupar todo el espacio
                                c.save()
                                
                                # Limpiar archivos temporales
                                os.remove(temp_image_path)
                                os.remove(file_path_temp)
                                
                                altura_pCourier5adas = pdf_height / 72
                                print(f"Páginas 3-{total_paginas} unificadas en una hoja de 6x{altura_pCourier5adas:.1f}\"")
                                log_text.insert(tk.END, f"✅ Páginas unificadas (6x{altura_pCourier5adas:.1f}\")\n")
                            else:
                                # No hay páginas para unir, usar el PDF completo
                                os.rename(file_path_temp, file_path)
                                
                        elif total_paginas == 2:
                            # Si solo hay 2 páginas, convertir la segunda a PDF de 6" de ancho
                            print(f"PDF tiene 2 páginas, procesando la página 2")
                            log_text.insert(tk.END, f"PDF tiene 2 páginas, usando la segunda\n")
                            
                            page = doc[1]
                            mat = fitz.Matrix(4, 4)  # Mayor resolución para impresión
                            pix = page.get_pixmap(matrix=mat)
                            img_data = pix.tobytes("png")
                            segunda_pagina = PILImage.open(io.BytesIO(img_data))
                            doc.close()
                            
                            # Recortar bordes blancos
                            segunda_pagina = recortar_bordes_blancos(segunda_pagina)
                            
                            # Redimensionar al ancho objetivo (300 DPI para impresión)
                            target_width = int(pdf_width / 72 * 300)
                            aspect_ratio = segunda_pagina.height / float(segunda_pagina.width)
                            new_height = int(target_width * aspect_ratio)
                            segunda_pagina = segunda_pagina.resize((target_width, new_height), PILImage.Resampling.LANCZOS)
                            
                            # Guardar como imagen temporal
                            temp_image_path = os.path.join(DOWNLOAD_DIR, f"temp_page2_{so}.png")
                            segunda_pagina.save(temp_image_path, 'PNG')
                            
                            # Calcular altura proporcional
                            img_aspect = new_height / float(target_width)
                            pdf_height = pdf_width * img_aspect
                            
                            # Crear PDF de 6" ancho con altura proporcional
                            c = canvas.Canvas(file_path, pagesize=(pdf_width, pdf_height))
                            c.drawImage(temp_image_path, 0, 0, 
                                       width=pdf_width, height=pdf_height,
                                       preserveAspectRatio=False)
                            c.save()
                            
                            os.remove(temp_image_path)
                            os.remove(file_path_temp)
                            
                        else:
                            # Si solo hay 1 página
                            print(f"PDF tiene solo 1 página")
                            log_text.insert(tk.END, f"PDF tiene solo 1 página\n")
                            
                            page = doc[0]
                            mat = fitz.Matrix(4, 4)  # Mayor resolución para impresión
                            pix = page.get_pixmap(matrix=mat)
                            img_data = pix.tobytes("png")
                            primera_pagina = PILImage.open(io.BytesIO(img_data))
                            doc.close()
                            
                            # Recortar bordes blancos
                            primera_pagina = recortar_bordes_blancos(primera_pagina)
                            
                            # Redimensionar al ancho objetivo (300 DPI para impresión)
                            target_width = int(pdf_width / 72 * 300)
                            aspect_ratio = primera_pagina.height / float(primera_pagina.width)
                            new_height = int(target_width * aspect_ratio)
                            primera_pagina = primera_pagina.resize((target_width, new_height), PILImage.Resampling.LANCZOS)
                            
                            temp_image_path = os.path.join(DOWNLOAD_DIR, f"temp_page1_{so}.png")
                            primera_pagina.save(temp_image_path, 'PNG')
                            
                            # Calcular altura proporcional
                            img_aspect = new_height / float(target_width)
                            pdf_height = pdf_width * img_aspect
                            
                            c = canvas.Canvas(file_path, pagesize=(pdf_width, pdf_height))
                            c.drawImage(temp_image_path, 0, 0, 
                                       width=pdf_width, height=pdf_height,
                                       preserveAspectRatio=False)
                            c.save()
                            
                            os.remove(temp_image_path)
                            os.remove(file_path_temp)
                        
                    except Exception as e:
                        print(f"Error procesando páginas: {e}")
                        log_text.insert(tk.END, f"Error procesando páginas: {e}\n")
                        # Si falla el procesamiento, usar el PDF completo
                        if os.path.exists(file_path_temp):
                            os.rename(file_path_temp, file_path)
                    
                    print(f"PDF generado: {pdf_filename}")
                    log_text.insert(tk.END, f"PDF generado: {pdf_filename}\n")
                    
                    # Cerrar la pestaña y volver a la principal
                    try:
                        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {}})
                    except Exception:
                        pass
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

                except Exception as e:
                    print(f"Error generando PDF: {e}")
                    log_text.insert(tk.END, f"Error generando PDF: {e}\n")
                    
                    # Asegurarse de volver a la ventana principal
                    try:
                        if len(driver.window_handles) > 1:
                            try:
                                driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {}})
                            except Exception:
                                pass
                            driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    except:
                        pass
                    
                    error_msg = f'Error: {str(e)}'
                    sheet.update_cell(fila_num, 2, error_msg)
                    # Agregar al historial con error
                    agregar_a_historial(so, error_msg, ref_oms, courier)
                    continue

                # Subir a Drive
                try:
                    file_metadata = {
                        'name': f"Factura_{ref_oms}_{so}.pdf",
                        'parents': [ROOT_FOLDER_ID]
                    }

                    media = MediaFileUpload(file_path, mimetype='application/pdf')
                    uploaded_file = service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id',
                        supportsAllDrives=True
                    ).execute()

                    file_id = uploaded_file.get('id')
                    file_link = f"https://drive.google.com/file/d/{file_id}/view"
                    
                    print(f"Archivo subido a Drive: {file_link}")
                    log_text.insert(tk.END, f"Drive link: {file_link}\n")

                    # Eliminar el archivo local después de subirlo
                    try:
                        os.remove(file_path)
                        print(f"Archivo local eliminado: {pdf_filename}")
                    except Exception as e:
                        print(f"No se pudo eliminar archivo local: {e}")

                    # Escribir el link en la columna B
                    sheet.update_cell(fila_num, 2, file_link)
                    # Agregar al historial exitosamente
                    agregar_a_historial(so, file_link, ref_oms, courier)
                    log_text.insert(tk.END, f"✅ SO {so} procesado exitosamente\n")

                except Exception as e:
                    print(f"Error subiendo a Drive: {e}")
                    log_text.insert(tk.END, f"Error subiendo a Drive: {e}\n")
                    
                    error_msg = f'Error subiendo: {str(e)}'
                    sheet.update_cell(fila_num, 2, error_msg)
                    # Agregar al historial con error
                    agregar_a_historial(so, error_msg, ref_oms, courier)
                    continue

            except Exception as e:
                print(f"Error procesando SO {so}: {e}")
                log_text.insert(tk.END, f"Error procesando SO {so}: {e}\n")
                
                registrar_error(
                    error_message=f"Error: {str(e)}",
                    agente=agente_inconv_var.get(),
                    order_id=so,
                    exception=e
                )
                
                error_msg = f'Error: {str(e)}'
                try:
                    sheet.update_cell(fila_num, 2, error_msg)
                    # Intentar agregar al historial aunque haya error
                    try:
                        ref_oms = so  # Si no se pudo obtener ref, usar SO
                        agregar_a_historial(so, error_msg, ref_oms, courier)
                    except:
                        pass
                except:
                    pass
                continue

    except Exception as e:
        print("Error general en facturas_oms_Courier4:", str(e))
        log_text.insert(tk.END, f"Error general: {e}\n")
        
        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )

    finally:
        log_text.insert(tk.END, "Proceso finalizado.\n")
        if driver:
            driver.quit()
        
        registrar_accion(
            total_orders=len(filas_a_procesar),
            agente=agente_inconv_var.get(),
        )

def asignar_Courier6_thread():
    global select_option_var, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find

    selected_column = action_var.get()
    ROOT_FOLDER_ID = "1TkpPijIs9ZB4DoylR4YtR6URpA1-vqGu" # Ajusta este ID según tu carpeta
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")

    spreadsheet_id = "1WjoEORNcH_S54N2o-zbRj1JrTZgM4hzFAKNIGjUVxS8"
    tanda = lote.get()
    print(tanda)
    if tanda == "1":
        changes_notas = cliente.open_by_key(spreadsheet_id).worksheet("Accionar Bot 1")

    elif tanda == "2":
        changes_notas = cliente.open_by_key(spreadsheet_id).worksheet("Accionar Bot 2")

    historial = cliente.open_by_key(spreadsheet_id).worksheet("Historial")

    # Obtener valores de la columna 1, excluyendo el encabezado
    valores_columna1 = changes_notas.col_values(1)[1:]
    valores_no_vacios_columna1 = [valor for valor in valores_columna1 if valor]

    # Determinar la cantidad de valores no vacíos en la columna 1
    cantidad_valores = len(valores_no_vacios_columna1)

    # Obtener la misma cantidad de valores de la columna 2, excluyendo el encabezado
    valores_columna2 = changes_notas.col_values(2)[1:cantidad_valores + 1]

    # Unir los valores de la columna 1 y 2 separados por comas
    valores_coma_separada_columna1 = ','.join(valores_no_vacios_columna1)
    valores_coma_separada_columna2 = ','.join(valores_columna2)

    print(f"Valores obtenidos de la columna 1: {valores_coma_separada_columna1}")
    print(f"Valores obtenidos de la columna 2: {valores_coma_separada_columna2}")

    try:
        lista_valores_columna1 = valores_coma_separada_columna1.split(',')

        log_text.insert(tk.END,
                        f"SO obtenidos de la columna {selected_column}: {valores_coma_separada_columna1}\n")

        lista_valores_columna2 = valores_coma_separada_columna2.split(',')

        log_text.insert(tk.END,
                        f"ref obtenidos de la columna {selected_column}: {valores_coma_separada_columna2}\n")

        # Configuración de Selenium
        chrome_options = Options()
        selenium_service = None
        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(
                service=Service(driver_path),
                options=chrome_options
            )
        except Exception:
            driver = webdriver.Chrome(
                service=selenium_service,
                options=chrome_options
            )
        set_company_custom_header(driver)

        # Página web oms
        driver.get("https://use1.omsapp.com/admin_login.php?clients_id=company")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email_address']")))
        username.clear()
        username.send_keys(usuario)
        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password.clear()
        password.send_keys(contrasena)
        button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, 'submit-admin'))).click()

        fila_vacia_resultados = 2

        for so, ref in zip(lista_valores_columna1, lista_valores_columna2):
            columna_a_test = historial.col_values(1)
            fila_vacia_historial = len(columna_a_test) + 1
            file_link = []
            files_link = []

            try:
                home_dir = os.path.expanduser("~")
                downloads_folder_en = os.path.join(home_dir, "Downloads")
                downloads_folder_es = os.path.join(home_dir, "Descargas")

                if os.path.isdir(downloads_folder_en):
                    downloads_folder = downloads_folder_en
                elif os.path.isdir(downloads_folder_es):
                    downloads_folder = downloads_folder_es
                else:
                    raise Exception("No se encontró la carpeta de descargas.")

                try:
                    import platform
                    if platform.system() == "Windows":
                        CSIDL_DESKTOPDIRECTORY = 0x10
                        buf = ctypes.create_unicode_buffer(wt.MAX_PATH)
                        if ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf) == 0:
                            p = buf.value
                            if os.path.isdir(p):
                                return p
                except Exception:
                    pass
                # Verificar si hay múltiples etiquetas generadas
                etiquetas_generadas = [f for f in os.listdir(downloads_folder) if
                                       f.startswith(ref) and f.endswith('.pdf')]

                for idx, etiqueta in enumerate(etiquetas_generadas, start=1):
                    # Si hay una sola etiqueta, usar el nombre original
                    if len(etiquetas_generadas) == 1:
                        file_name = f"{ref}.pdf"
                    else:
                        # Si hay múltiples etiquetas, agregar un sufijo numérico
                        file_name = f"{ref}-{idx}.pdf"

                file_path = os.path.join(downloads_folder, etiqueta)

                # Subir el archivo a Google Drive
                file_metadata = {
                    'name': file_name,
                    'parents': [ROOT_FOLDER_ID]
                }

                media = MediaFileUpload(file_path, mimetype='application/pdf')
                uploaded_file = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()

                file_id = uploaded_file.get('id')
                file_link = f"https://drive.google.com/file/d/{file_id}/view"
                files_link.append(file_link)

                print(f'Archivo subido exitosamente. ID: {file_id}')

                changes_notas.update_cell(fila_vacia_resultados, 4, f'{file_link}')
                historial.update_cell(fila_vacia_historial, 4, f'{file_link}')

                if file_link:
                    try:
                        time.sleep(2)
                        # Entrar a PO y accionar
                        invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                        driver.get(invoice_url)

                        log_text.insert(tk.END, f"Trabajando en orden:{so}\n")
                        print(f"Trabajando en orden:{so}")

                        button_Customs = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                        button_Customs.click()

                        button_Changes = WebDriverWait(button_Customs, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Changes')]")))
                        button_Changes.click()

                        textarea = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.NAME, "PCF_CHANGES")))


                        resultado_notas = f"Orden viaja por Courier6  \nLink de la etiqueta: {file_link}\n"

                        # Obtener el contenido existente del textarea
                        existing_text = textarea.get_attribute('value')
                        textarea.clear()
                        notas_pegar = f"\n{file_link}\n"

                        # Añadir el nuevo texto al contenido existente
                        new_text = existing_text + notas_pegar
                        textarea.send_keys(new_text)

                        # Guardar
                        button_save_so = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()
                        time.sleep(2)

                        # Pegar las notas
                        button_notas = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//span[contains(text(), 'Notes and payment history')]")))
                        button_notas.click()

                        textarea = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.NAME, "note_text")))
                        textarea.clear()

                        textarea.send_keys(resultado_notas)

                        # Guardar progreso
                        button_save_clon = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                        changes_notas.update_cell(fila_vacia_resultados, 3, f'Ok')
                        historial.update_cell(fila_vacia_historial, 3, f'Ok')

                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                        fila_vacia_resultados += 1

                    except Exception as e:
                        print(f"Error al realizar orden {so}", str(e))
                        log_text.insert(tk.END, f"Error al realizar orden {so}\n")

                        changes_notas.update_cell(fila_vacia_resultados, 3, f'Error en oms')
                        historial.update_cell(fila_vacia_historial, 3, f'Error en oms')

                        historial.update_cell(fila_vacia_historial, 1, f'{so}')
                        historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                        historial.update_cell(fila_vacia_historial, 5, fecha_actual)
                        fila_vacia_resultados += 1

                else:
                    changes_notas.update_cell(fila_vacia_resultados, 3, f'Error, no se encontro el pdf')
                    historial.update_cell(fila_vacia_historial, 3, f'Error, no se encontro el pdf')

                    historial.update_cell(fila_vacia_historial, 1, f'{so}')
                    historial.update_cell(fila_vacia_historial, 2, f'{ref}')
                    historial.update_cell(fila_vacia_historial, 5, fecha_actual)

                    fila_vacia_resultados += 1

            except Exception as e:
                print(f"Error al procesar SO {so}, no se encuentra el pdf: {str(e)}")
                log_text.insert(tk.END, f"Error al procesar SO {so}, no se encuentra el pdf: {str(e)}\n")
                
                changes_notas.update_cell(fila_vacia_resultados, 3, f'Error al encontrar PDF')

                historial.update_cell(fila_vacia_historial, 3, f'Error al encontrar PDF')
                historial.update_cell(fila_vacia_historial, 1, f'{so}')
                historial.update_cell(fila_vacia_historial, 2, f'{ref}')

                fila_vacia_resultados += 1
                continue

    except Exception as e:
        print("Error al realizar acción", str(e))
        log_text.insert(tk.END, f"Error al realizar acción\n")

    finally:
        # Mostrar mensaje de éxito
        log_text.insert(tk.END, f"Finalizado\n")
        driver.quit()

def clasificar_productos_con_ia(rows, pais, courier, sheets_service=None, sheet_id=None, sheet_name=None, start_row=2, col='L'):
    """
    Clasifica productos usando la API de Gemini con procesamiento concurrente.
    rows: lista de tuplas con los datos de Redshift
    sheets_service: servicio de Google Sheets para guardar progreso incremental
    sheet_id: ID de la hoja de destino
    sheet_name: nombre de la pestaña
    start_row: fila inicial donde empezar a pegar (por defecto 2)
    Retorna: lista de clasificaciones
    """
    import hashlib
    import json
    import requests
    import threading
    import time
    import difflib
    import datetime
    import importlib
    from concurrent.futures import ThreadPoolExecutor, as_completed
    try:
        genai = importlib.import_module('google.genai')
        genai_types = importlib.import_module('google.genai.types')
    except Exception:
        genai = None
        genai_types = None

    if not rows:
        return []

    tiempo_inicio = time.time()

    # Defaults afinados (Ola 1):
    # - gemini-2.5-flash: modelo final estable, mejor calidad que gemini-3-flash-preview
    # - batch_size 6: menos confusión por items mezclados en un mismo prompt
    # - max_workers 6: compensa el batch más chico para mantener throughput
    MODEL_NAME = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')
    USE_VERTEX = os.getenv('GEMINI_USE_VERTEX', '0').strip().lower() in ('1', 'true', 'yes', 'si')
    API_KEY = os.getenv('GEMINI_API_KEY', '')

    # Config Fase 1 (prioriza throughput sin perder calidad)
    max_workers = int(os.getenv('TAXES_IA_MAX_WORKERS', '6'))
    batch_size = int(os.getenv('TAXES_IA_BATCH_SIZE', '6'))
    max_retries = int(os.getenv('TAXES_IA_MAX_RETRIES', '4'))
    timeout_sec = int(os.getenv('TAXES_IA_TIMEOUT', '45'))
    save_every = int(os.getenv('TAXES_IA_SAVE_EVERY', '80'))
    min_pause = float(os.getenv('TAXES_IA_MIN_PAUSE', '0.05'))
    max_review_items = int(os.getenv('TAXES_IA_MAX_REVIEW_ITEMS', '120'))
    # Por defecto guardamos también categorías genéricas para que el cache siempre aprenda y sea reutilizable.
    cache_include_los_demas = os.getenv('TAXES_CACHE_INCLUDE_LOS_DEMAS', '1').strip().lower() in ('1', 'true', 'yes', 'si')

    # Cache en Google Sheets (centralizado por defecto para reusar entre corridas/couriers)
    cache_sheet_id = os.getenv('TAXES_CACHE_SHEET_ID', '').strip() or '14lnY-azLXdPIX4BEVarNN7chFu6SmHbSECrB6stkCzc'
    cache_sheet_name = os.getenv('TAXES_CACHE_SHEET_NAME', 'CACHE IA').strip() or 'CACHE IA'

    # Debug/diagnóstico: forzar escritura en cache si se necesita investigar
    FORCE_CACHE_WRITE = os.getenv('TAXES_CACHE_FORCE_WRITE', '0').strip().lower() in ('1', 'true', 'yes', 'si')

    if USE_VERTEX:
        print('[clasificar_ia] Modo Vertex habilitado (si falla token, hace fallback a Gemini API key).')
    sin_modelo_ia = not API_KEY and not USE_VERTEX
    if sin_modelo_ia:
        print('[clasificar_ia] No se encontró GEMINI_API_KEY y Vertex no está habilitado.')

    gemini_client = None
    if API_KEY and genai is not None:
        try:
            gemini_client = genai.Client(api_key=API_KEY)
        except Exception as e:
            print(f"[clasificar_ia] No se pudo inicializar SDK de Gemini: {e}")

    # Obtener categorías desde Google Sheets, manteniendo independencia por courier.
    try:
        sheets_service_categorias = sheets_service
        if sheets_service_categorias is None:
            from google.oauth2.service_account import Credentials
            from googleapiclient.discovery import build

            cred_file = 'credenciales_drive.json'
            scopes = ['https://www.googleapis.com/auth/spreadsheets']
            credentials = Credentials.from_service_account_file(cred_file, scopes=scopes)
            sheets_service_categorias = build('sheets', 'v4', credentials=credentials)

        def _force_range_from_row_2(range_a1):
            txt = str(range_a1 or '').strip()
            if not txt:
                return txt

            # Asegura lectura desde fila 2+ para no tomar encabezados.
            m = re.match(
                r'^(?P<prefix>[^!]+!)?(?P<start_col>[A-Za-z]+)(?P<start_row>\d+)?(?::(?P<end_col>[A-Za-z]+)?(?P<end_row>\d+)?)?$',
                txt
            )
            if not m:
                return txt

            prefix = m.group('prefix') or ''
            start_col = m.group('start_col')
            start_row_raw = m.group('start_row')
            end_col = m.group('end_col') or ''
            end_row_raw = m.group('end_row')

            start_row = int(start_row_raw) if start_row_raw else 2
            if start_row < 2:
                start_row = 2

            if ':' in txt:
                right = f"{end_col}{end_row_raw or ''}"
                if not right:
                    right = end_col or ''
                return f"{prefix}{start_col}{start_row}:{right}"
            return f"{prefix}{start_col}{start_row}"

        def _canon_simple(v):
            txt = ' '.join(str(v or '').strip().lower().split())
            return ''.join(
                ch for ch in unicodedata.normalize('NFD', txt)
                if unicodedata.category(ch) != 'Mn'
            )

        def _try_get_categorias(spreadsheet_id, ranges_to_try):
            if not spreadsheet_id:
                return []
            for r in ranges_to_try:
                rr = _force_range_from_row_2(r)
                if not rr:
                    continue
                try:
                    result = sheets_service_categorias.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=rr
                    ).execute()
                    encabezados_invalidos = {
                        'categoria',
                        'categorias',
                        'taxes',
                        'tabla taxes',
                    }
                    vals = [
                        row[0].strip()
                        for row in result.get('values', [])
                        if row
                        and str(row[0]).strip()
                        and not str(row[0]).strip().isdigit()
                        and _canon_simple(str(row[0]).strip()) not in encabezados_invalidos
                    ]
                    if vals:
                        print(f"[clasificar_ia] Categorías cargadas desde {spreadsheet_id} | {rr}: {len(vals)}")
                        return vals
                except Exception:
                    continue
            return []

        courier_key = str(courier or '').strip().upper().replace(' ', '_')
        default_ranges_by_courier = {
            'Courier4': 'TABLA Taxes!B2:B',
            'Courier1': 'TABLA Taxes!A2:A',
            'Courier_5': 'TABLA Taxes!A2:A',
            'Courier3': 'TABLA Taxes!A2:A',
        }
        default_ranges_cfg = default_ranges_by_courier.get(courier_key, 'TABLA Taxes!A2:A')

        # Prioridad de configuración:
        # 1) TAXES_CATEGORIAS_RANGES_<COURIER> (ej: TAXES_CATEGORIAS_RANGES_Courier4)
        # 2) TAXES_CATEGORIAS_RANGES (global)
        # 3) default por courier
        ranges_cfg = (
            os.getenv(f'TAXES_CATEGORIAS_RANGES_{courier_key}', '').strip()
            or os.getenv('TAXES_CATEGORIAS_RANGES', '').strip()
            or default_ranges_cfg
        )
        print(f"[clasificar_ia] Rango de categorías para {courier}: {ranges_cfg}")
        ranges_to_try = [_force_range_from_row_2(x) for x in ranges_cfg.split(',') if x.strip()]

        # 1) Primero: sheet del courier actual (independiente)
        categorias = _try_get_categorias(sheet_id, ranges_to_try)

        # No usar fallback cruzado entre couriers: siempre se toma la tabla del mismo sheet.
        categorias = list(dict.fromkeys([str(c).strip() for c in categorias if str(c).strip()]))

        categorias_genericas = {
            'los demas',
            'los demas con impuestos internos',
            'los demas con impuesto interno',
            'otros',
            'otras',
            'otros productos',
            'others',
            'other',
        }
        categoria_generica = next(
            (c for c in categorias if _canon_simple(c) in categorias_genericas),
            None
        )

        # Evitar sesgo al primer valor (ej. 'Birds'): para Courier4 preferir 'Others' como fallback.
        if courier_key == 'Courier4':
            categoria_fallback = (
                categoria_generica
                or next((c for c in categorias if _canon_simple(c) in {'others', 'other', 'otros', 'otras'}), None)
                or (categorias[0] if categorias else 'SIN_CATEGORIA')
            )
        else:
            categoria_fallback = categoria_generica or (categorias[0] if categorias else 'SIN_CATEGORIA')

        categorias_set = set(categorias)
        categorias_lower_map = {c.lower(): c for c in categorias}
    except Exception as e:
        print(f'[clasificar_ia] Error obteniendo categorías: {e}')
        return ['SIN_CATEGORIA'] * len(rows)

    if not categorias:
        return ['SIN_CATEGORIA'] * len(rows)

    if sin_modelo_ia:
        return [categoria_fallback] * len(rows)

    # Servicios auxiliares
    def _normalizar_texto(v):
        return ' '.join(str(v or '').strip().lower().split())

    def _canon(v):
        txt = _normalizar_texto(v)
        return ''.join(
            ch for ch in unicodedata.normalize('NFD', txt)
            if unicodedata.category(ch) != 'Mn'
        )

    def _tokenizar(v):
        tokens = re.findall(r'[a-z0-9]+', _canon(v))
        stopwords = {
            'de', 'del', 'la', 'las', 'el', 'los', 'y', 'con', 'sin', 'para', 'por', 'en',
            'the', 'and', 'with', 'for', 'from', 'pack', 'size', 'color', 'inch', 'male',
            'female', 'black', 'white'
        }
        return [t for t in tokens if len(t) > 2 and t not in stopwords]

    def _parse_categoria_json(categoria_raw):
        txt = str(categoria_raw or '').strip()
        if not txt:
            return None

        candidatos = [txt]

        # Caso frecuente: JSON con comillas duplicadas {""k"":...}
        if '""' in txt:
            candidatos.append(txt.replace('""', '"'))

        # Caso doble-serializado: "{...}"
        if txt.startswith('"') and txt.endswith('"'):
            inner = txt[1:-1].replace('\\"', '"')
            candidatos.append(inner)

        for cand in candidatos:
            try:
                data = json.loads(cand)
                if isinstance(data, str):
                    data = json.loads(data)
                if isinstance(data, dict):
                    return data
            except Exception:
                continue
        return None

    def _extraer_datos_producto(row):
        categoria_json = row[2] if len(row) > 2 else ''
        nombre_producto = row[3] if len(row) > 3 else ''
        sku = str(row[1] if len(row) > 1 else '').strip().upper()
        categoria_madre = ''
        detalles = ''
        categoria_original = ''
        try:
            data = _parse_categoria_json(categoria_json)
            if isinstance(data, dict):
                detalles_parts = []
                for attr in data.get('attributes', []):
                    if not isinstance(attr, dict):
                        continue
                    attr_name = str(attr.get('name', '')).strip()
                    attr_value = str(attr.get('value', '')).strip()
                    if not attr_value:
                        continue
                    if attr_name == 'descripcion_categoria_aduanera' and not categoria_madre:
                        categoria_madre = attr_value
                    detalles_parts.append(f"{attr_name}: {attr_value}" if attr_name else attr_value)
                if 'categoryTree' in data and isinstance(data['categoryTree'], list):
                    tree_txt = ' > '.join([str(x) for x in data['categoryTree']])
                    tree_txt = tree_txt.replace('Others', '').replace(' >  > ', ' > ').strip()
                    if tree_txt:
                        detalles_parts.append(tree_txt)
                detalles = ' | '.join([p for p in detalles_parts if p])[:800]
            categoria_original = categoria_madre or detalles or str(categoria_json or '').strip()
        except Exception:
            detalles = str(categoria_json)
            categoria_original = str(categoria_json or '').strip()

        # Evitar celdas enormes en cache de Sheets
        categoria_original = categoria_original[:500]
        return categoria_madre, detalles, str(nombre_producto or '').strip(), sku, categoria_original

    def _cache_key_hash(categoria_madre, detalles, nombre, sku=''):
        base = (
            f"{_normalizar_texto(categoria_madre)}|{_normalizar_texto(detalles)}|"
            f"{_normalizar_texto(nombre)}|{_normalizar_texto(sku)}"
        )
        return hashlib.sha1(base.encode('utf-8')).hexdigest()

    def _parse_json_array(text):
        raw = (text or '').strip()
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            pass
        start = raw.find('[')
        end = raw.rfind(']')
        if start != -1 and end != -1 and end > start:
            chunk = raw[start:end + 1]
            try:
                parsed = json.loads(chunk)
                if isinstance(parsed, list):
                    return parsed
            except Exception:
                pass
        return []

    def _normalizar_categoria(cat):
        c = str(cat or '').strip().replace('"', '').replace("'", '')
        if not c:
            return categoria_fallback
        if c in categorias_set:
            return c
        cl = c.lower()
        if cl in categorias_lower_map:
            return categorias_lower_map[cl]

        # Match flexible por forma canónica (sin acentos, espacios normalizados)
        canon_c = _canon(c)
        for original in categorias:
            if _canon(original) == canon_c:
                return original

        # Match por contención: solo aceptamos cuando la categoría oficial es
        # razonablemente larga (>= 4 chars) o tiene varias palabras. Antes
        # categorías cortas como "PC", "Mac", "TV" disparaban falsos positivos
        # porque dos letras aparecen como substring en muchos textos.
        for original in categorias:
            canon_o = _canon(original)
            if not canon_o:
                continue
            tiene_varias_palabras = ' ' in canon_o
            if not tiene_varias_palabras and len(canon_o) < 4:
                continue
            if canon_o in canon_c or canon_c in canon_o:
                return original

        # Fuzzy match conservador para variaciones de redacción
        best_cat = None
        best_score = 0.0
        tokens_c = set(_tokenizar(c))
        for original in categorias:
            ratio = difflib.SequenceMatcher(None, _canon(c), _canon(original)).ratio()
            tokens_o = set(_tokenizar(original))
            if tokens_c and tokens_o:
                jacc = len(tokens_c.intersection(tokens_o)) / max(len(tokens_c.union(tokens_o)), 1)
            else:
                jacc = 0.0
            score = max(ratio, jacc)
            if score > best_score:
                best_score = score
                best_cat = original

        if best_cat and best_score >= 0.86:
            return best_cat

        return categoria_fallback

    def _es_categoria_generica(cat):
        canon = _canon(cat)
        return canon in categorias_genericas

    def _mejor_match_lexico(query_text):
        q_tokens = set(_tokenizar(query_text))
        query_canon = _canon(query_text)
        if not q_tokens and not query_canon:
            return None, 0.0, 0.0

        top_1 = ('', 0.0)
        top_2 = ('', 0.0)
        for cat in categorias:
            cat_canon = _canon(cat)
            c_tokens = set(_tokenizar(cat))
            ratio = difflib.SequenceMatcher(None, query_canon, cat_canon).ratio() if query_canon and cat_canon else 0.0
            jacc = 0.0
            if q_tokens and c_tokens:
                jacc = len(q_tokens.intersection(c_tokens)) / max(len(q_tokens.union(c_tokens)), 1)

            score = (0.55 * jacc) + (0.45 * ratio)
            if _es_categoria_generica(cat):
                score -= 0.20

            if score > top_1[1]:
                top_2 = top_1
                top_1 = (cat, score)
            elif score > top_2[1]:
                top_2 = (cat, score)

        margin = top_1[1] - top_2[1]
        return top_1[0] or None, top_1[1], margin

    def _match_categoria_literal(query_text):
        """Busca coincidencias literales contra categorías específicas antes de usar scoring.

        Esto evita que descripciones con señales claras, como 'Fragrances', caigan en una
        categoría genérica cuando la categoría oficial ya aparece explícitamente en la evidencia.

        Importante: las categorías muy cortas (PC, Mac, TV) generan falsos
        positivos brutales por substring, así que las saltamos aquí. Sólo
        aceptamos categorías de 4+ chars o con varias palabras.
        """
        query_canon = _canon(query_text)
        if not query_canon:
            return None

        for original in categorias:
            if _es_categoria_generica(original):
                continue
            canon_original = _canon(original)
            if not canon_original:
                continue
            tiene_varias_palabras = ' ' in canon_original
            if not tiene_varias_palabras and len(canon_original) < 4:
                continue
            if canon_original in query_canon:
                return original
        return None

    def _resolver_los_demas_por_evidencia(item):
        query_text = ' '.join([
            str(item.get('madre', '')).strip(),
            str(item.get('detalles', '')).strip(),
            str(item.get('nombre', '')).strip(),
            str(item.get('sku', '')).strip(),
        ]).strip()
        literal_cat = _match_categoria_literal(query_text)
        if literal_cat:
            return literal_cat
        best_cat, best_score, margin = _mejor_match_lexico(query_text)
        if not best_cat:
            return ''
        # Solo corrige cuando la evidencia textual es claramente mejor que una categoría genérica.
        if best_score >= 0.55 and margin >= 0.10 and not _es_categoria_generica(best_cat):
            return best_cat
        return ''

    def _top_categorias_candidatas(madre, detalles, nombre, top_k=8):
        """Devuelve categorías candidatas por similitud léxica general, sin reglas puntuales.

        Si el producto no tiene tokens útiles, NO devolvemos las primeras N
        categorías de la lista (eso sesgaba todo a 'Pet Supplies, Birds, Cats,
        Dogs, ...'). En cambio, devolvemos una muestra DIVERSA tomando un
        ítem cada cierto paso a lo largo de la lista, lo que da al modelo
        un panorama amplio sin sesgarlo.
        """
        def _muestra_diversa(top_k):
            no_genericas = [c for c in categorias if not _es_categoria_generica(c)]
            if not no_genericas:
                return list(categorias)[:top_k]
            n = len(no_genericas)
            if n <= top_k:
                return no_genericas[:top_k]
            paso = max(1, n // top_k)
            return [no_genericas[i] for i in range(0, n, paso)][:top_k]

        query_blob = f"{madre} {detalles} {nombre}"
        q_tokens = set(_tokenizar(query_blob))
        if not q_tokens:
            return _muestra_diversa(top_k)

        scored = []
        for cat in categorias:
            if _es_categoria_generica(cat):
                continue
            c_tokens = set(_tokenizar(cat))
            if not c_tokens:
                continue
            inter = len(q_tokens.intersection(c_tokens))
            precision = inter / max(len(c_tokens), 1)
            recall = inter / max(len(q_tokens), 1)
            score = (0.65 * precision) + (0.35 * recall)
            scored.append((score, cat))

        scored.sort(key=lambda x: x[0], reverse=True)
        top = [cat for score, cat in scored[:top_k] if score > 0]
        if top:
            return top
        return _muestra_diversa(top_k)

    def _score_categoria_con_evidencia(query_text, categoria):
        query_canon = _canon(query_text)
        categoria_canon = _canon(categoria)
        if not query_canon or not categoria_canon:
            return 0.0

        q_tokens = set(_tokenizar(query_text))
        c_tokens = set(_tokenizar(categoria))
        ratio = difflib.SequenceMatcher(None, query_canon, categoria_canon).ratio()
        jacc = 0.0
        if q_tokens and c_tokens:
            jacc = len(q_tokens.intersection(c_tokens)) / max(len(q_tokens.union(c_tokens)), 1)

        score = (0.55 * jacc) + (0.45 * ratio)
        if _es_categoria_generica(categoria):
            score -= 0.20
        return score

    def _corregir_categoria_por_evidencia(item, categoria_asignada):
        query_text = ' '.join([
            str(item.get('madre', '')).strip(),
            str(item.get('detalles', '')).strip(),
            str(item.get('nombre', '')).strip(),
            str(item.get('sku', '')).strip(),
            str(item.get('categoria_original', '')).strip(),
        ]).strip()

        if not query_text:
            return categoria_asignada

        literal_cat = _match_categoria_literal(query_text)
        if literal_cat:
            return literal_cat

        best_cat, best_score, _ = _mejor_match_lexico(query_text)
        assigned_score = _score_categoria_con_evidencia(query_text, categoria_asignada)

        if categoria_generica and _canon(categoria_asignada) == _canon(categoria_generica):
            categoria_sugerida = _resolver_los_demas_por_evidencia(item)
            if categoria_sugerida:
                return categoria_sugerida

        # Más estricto (Ola 2a): solo reemplazamos la asignación de la IA cuando la
        # evidencia léxica es claramente más fuerte. Antes con scores apenas mejores
        # se sobreescribían asignaciones correctas y empujábamos a categorías genéricas.
        if best_cat and best_score >= 0.50 and (
            assigned_score < 0.20 or best_score >= assigned_score + 0.15
        ):
            return best_cat

        return categoria_asignada

    # Contenedor de ejemplos verificados (Ola 2c). Se llena después de cargar el
    # cache y los closures de los prompt builders lo capturan por referencia.
    verified_examples = []

    def _format_verified_examples(max_examples=6):
        if not verified_examples:
            return ""
        lines = []
        for ex in verified_examples[:max_examples]:
            cat_orig = str(ex.get('cat_original', '')).strip()
            sku = str(ex.get('sku', '')).strip()
            categoria = str(ex.get('categoria', '')).strip()
            if not categoria:
                continue
            # Truncar campos largos para no inflar el prompt
            cat_orig_short = (cat_orig[:240] + '…') if len(cat_orig) > 240 else cat_orig
            lines.append(
                f"- categoria_original: {cat_orig_short} | sku: {sku} -> categoria: {categoria}"
            )
        if not lines:
            return ""
        return (
            "EJEMPLOS_VERIFICADOS (clasificaciones confirmadas como correctas; usalos como guia de estilo):\n"
            + "\n".join(lines)
            + "\n\n"
        )

    def _build_batch_prompt(items_batch):
        items_text = []
        candidatos_union = []
        for item in items_batch:
            candidatas = item.get('candidatas', [])
            candidatos_union.extend(candidatas)
            items_text.append(
                f"- id: {item['id']} | madre: {item['madre']} | detalles: {item['detalles']} | nombre: {item['nombre']}"
                f" | categoria_original: {item.get('categoria_original', '')}"
                f" | candidatas_sugeridas: {json.dumps(candidatas, ensure_ascii=False)}"
            )

        # IMPORTANTE: mostramos TODAS las categorías oficiales al modelo. El cap
        # de 80 que había excluía categorías clave (Fragrances pos 207, Shoes
        # 165, Watches 167, Health & Beauty 196) y el modelo, al no verlas,
        # anclaba en lo que tenía a mano (PC, Mac, etc.). El env var
        # TAXES_IA_MAX_CATEGORIAS_EN_PROMPT permite recortar si en algún
        # courier la lista crece demasiado para el contexto del modelo.
        max_cats = int(os.getenv('TAXES_IA_MAX_CATEGORIAS_EN_PROMPT', '0') or 0)
        lista_candidatas_global = []
        seen = set()
        for c in candidatos_union + categorias:
            cs = str(c).strip()
            if not cs or cs in seen:
                continue
            seen.add(cs)
            lista_candidatas_global.append(cs)
            if max_cats and len(lista_candidatas_global) >= max_cats:
                break
        regla_generica = (
            f"4) Usa '{categoria_generica}' solo si ninguna categoria de la lista aplica claramente.\n"
            if categoria_generica else
            "4) Si ninguna coincide exacto, elige la categoria oficial mas cercana por semantica; nunca inventes categorias.\n"
        )

        ejemplos_block = _format_verified_examples(max_examples=6)

        return (
            f"Eres un experto clasificador aduanero para {pais}. "
            f"Debes clasificar cada producto en una categoria EXACTA de la lista oficial.\n\n"
            "REGLAS:\n"
            "1) Solo puedes responder categorias existentes en LISTA_CATEGORIAS.\n"
            "2) Prioriza la categoria mas especifica posible.\n"
            "3) Usa candidatas_sugeridas como primera referencia; si ninguna aplica, usa LISTA_CATEGORIAS completa.\n"
            + regla_generica +
            "5) Responde SOLO JSON valido (array), sin texto extra.\n"
            "6) Incluye un campo 'razon' breve (max 15 palabras) explicando por que elegiste esa categoria.\n"
            "7) Formato de salida: [{\"id\":\"...\",\"categoria\":\"...\",\"razon\":\"...\"}]\n\n"
            + ejemplos_block +
            f"LISTA_CATEGORIAS: {json.dumps(lista_candidatas_global, ensure_ascii=False)}\n\n"
            "PRODUCTOS:\n"
            + "\n".join(items_text)
        )

    def _build_single_prompt(item):
        candidatas = item.get('candidatas', [])
        # Igual que en batch: mostramos toda la lista oficial al modelo. Cap
        # opcional vía env var (compartido con el batch) para casos extremos.
        max_cats = int(os.getenv('TAXES_IA_MAX_CATEGORIAS_EN_PROMPT', '0') or 0)
        extendidas = []
        vistos = set()
        for c in candidatas + categorias:
            cs = str(c).strip()
            if not cs or cs in vistos:
                continue
            vistos.add(cs)
            extendidas.append(cs)
            if max_cats and len(extendidas) >= max_cats:
                break
        regla_generica = (
            f"Si ninguna categoria aplica de forma clara, usa '{categoria_generica}'.\n\n"
            if categoria_generica else
            "Si ninguna coincide exacto, elige la categoria oficial mas cercana por semantica; no inventes categorias.\n\n"
        )

        ejemplos_block = _format_verified_examples(max_examples=4)

        return (
            f"Clasifica 1 producto para {pais}.\n"
            "Responde SOLO JSON valido con este formato exacto: {\"categoria\":\"...\",\"razon\":\"...\"}\n"
            "Debe ser exactamente una categoria de LISTA_CATEGORIAS.\n"
            "Incluye 'razon' breve (max 15 palabras) explicando la elección.\n"
            "Usa primero la evidencia del producto: categoria original, nombre y SKU.\n"
            + regla_generica +
            ejemplos_block +
            f"LISTA_CATEGORIAS: {json.dumps(extendidas, ensure_ascii=False)}\n\n"
            f"PRODUCTO: madre={item.get('madre','')} | detalles={item.get('detalles','')} | nombre={item.get('nombre','')} | sku={item.get('sku','')} | categoria_original={item.get('categoria_original','')}"
        )

    def _parse_single_categoria(text):
        raw = (text or '').strip()
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict):
                return str(obj.get('categoria', '')).strip()
        except Exception:
            pass
        start = raw.find('{')
        end = raw.rfind('}')
        if start != -1 and end != -1 and end > start:
            chunk = raw[start:end + 1]
            try:
                obj = json.loads(chunk)
                if isinstance(obj, dict):
                    return str(obj.get('categoria', '')).strip()
            except Exception:
                pass
        return ''

    # Schemas con campo 'razon' opcional (Ola 2b): pedir una justificación corta
    # mejora la consistencia del modelo y nos da audit trail sin penalizar tiempo
    # (la razón es de pocos tokens). 'razon' no es obligatorio para no romper
    # respuestas que vengan sin él.
    batch_response_schema = {
        'type': 'array',
        'items': {
            'type': 'object',
            'properties': {
                'id': {'type': 'string'},
                'categoria': {'type': 'string'},
                'razon': {'type': 'string'},
            },
            'required': ['id', 'categoria'],
        },
    }

    single_response_schema = {
        'type': 'object',
        'properties': {
            'categoria': {'type': 'string'},
            'razon': {'type': 'string'},
        },
        'required': ['categoria'],
    }

    def _reclasificar_item(item):
        prompt = _build_single_prompt(item)
        for intento in range(max_retries):
            try:
                api_result = None
                if USE_VERTEX:
                    response = _call_vertex_api(prompt)
                    if response is not None:
                        if response.status_code == 429:
                            time.sleep(min((2 ** intento), 10))
                            continue
                        if response.status_code >= 500:
                            time.sleep(min((2 ** intento), 10))
                            continue
                        if response.status_code != 200:
                            return categoria_fallback, '', ''

                        body = response.json()
                        text = _extract_text_from_response(body)
                        cat_raw = _parse_single_categoria(text)
                        return _normalizar_categoria(cat_raw), cat_raw, ''

                api_result = _call_gemini_api(prompt, single_response_schema)

                if api_result.get('status_code', 0) == 429:
                    time.sleep(min((2 ** intento), 10))
                    continue
                if api_result.get('status_code', 0) >= 500:
                    time.sleep(min((2 ** intento), 10))
                    continue
                if api_result.get('status_code') != 200:
                    return categoria_fallback, '', ''

                parsed = api_result.get('parsed')
                razon = ''
                if isinstance(parsed, dict):
                    cat_raw = str(parsed.get('categoria', '')).strip()
                    razon = str(parsed.get('razon', '')).strip()
                else:
                    cat_raw = _parse_single_categoria(str(api_result.get('text', '') or ''))
                return _normalizar_categoria(cat_raw), cat_raw, razon
            except Exception:
                if intento < max_retries - 1:
                    time.sleep(min((2 ** intento), 10))
        return categoria_fallback, '', ''

    def _debe_revisar_los_demas(item, cat):
        if not categoria_generica:
            return False
        if _canon(cat) != _canon(categoria_generica):
            return False
        madre = str(item.get('madre', '')).strip()
        detalles = str(item.get('detalles', '')).strip()
        nombre = str(item.get('nombre', '')).strip()
        candidatas = [
            c for c in item.get('candidatas', [])
            if str(c).strip() and _canon(str(c).strip()) != _canon(categoria_generica)
        ]
        # Si hay info suficiente o sugerencias, no aceptar Los demás de primera
        return bool(madre or detalles or nombre or candidatas)

    # Modelo de fallback automático cuando el primario devuelve 500 persistentes.
    FALLBACK_MODEL = os.getenv('TAXES_IA_FALLBACK_MODEL', 'gemini-2.5-flash-lite').strip() or 'gemini-2.5-flash-lite'

    def _call_gemini_api(prompt, response_schema=None, model_override=None):
        model_to_use = model_override or MODEL_NAME
        # Camino preferido: SDK moderno de Gemini con salida estructurada.
        if gemini_client is not None and genai_types is not None:
            try:
                # thinking_budget=0 desactiva el thinking mode de gemini-2.5-flash.
                # Para clasificación determinística no aporta valor y causa respuestas
                # que mezclan texto de razonamiento con el JSON esperado.
                _thinking_cfg = None
                try:
                    _thinking_cfg = genai_types.ThinkingConfig(thinking_budget=0)
                except Exception:
                    pass

                cfg = genai_types.GenerateContentConfig(
                    temperature=0.0,
                    max_output_tokens=2048,
                    thinking_config=_thinking_cfg,
                )
                if response_schema is not None:
                    cfg = genai_types.GenerateContentConfig(
                        temperature=0.0,
                        max_output_tokens=2048,
                        response_mime_type='application/json',
                        response_schema=response_schema,
                        thinking_config=_thinking_cfg,
                    )

                response = gemini_client.models.generate_content(
                    model=model_to_use,
                    contents=prompt,
                    config=cfg,
                )

                return {
                    'status_code': 200,
                    'parsed': getattr(response, 'parsed', None),
                    'text': getattr(response, 'text', '') or '',
                    'model_used': model_to_use,
                }
            except Exception as e:
                msg = str(e)
                status = 500
                msg_upper = msg.upper()
                # Detectamos el código real reportado por la API. El orden
                # importa: chequeamos 400/403 explícitos antes que las
                # heurísticas de texto, porque mensajes como "API key expired"
                # vienen acompañados de "400 INVALID_ARGUMENT" y queremos que
                # el status mostrado coincida con el real.
                if '429' in msg or 'RESOURCE_EXHAUSTED' in msg_upper:
                    status = 429
                elif '400' in msg or 'INVALID_ARGUMENT' in msg_upper:
                    # 400 incluye API_KEY_INVALID, API key expired, schema inválido
                    status = 400
                elif '403' in msg or 'PERMISSION_DENIED' in msg_upper:
                    status = 403
                elif 'API KEY' in msg_upper:
                    # Heurística por texto cuando no vino el código.
                    status = 403
                return {
                    'status_code': status,
                    'parsed': None,
                    'text': '',
                    'error': msg,
                    'model_used': model_to_use,
                }

        # Fallback por compatibilidad: REST, intentando también forzar JSON schema.
        url = f'https://generativelanguage.googleapis.com/v1beta/models/{model_to_use}:generateContent?key={API_KEY}'
        generation_config = {'temperature': 0.0, 'maxOutputTokens': 2048}
        if response_schema is not None:
            generation_config.update({
                'responseMimeType': 'application/json',
                'responseSchema': response_schema,
            })
        payload = {
            'contents': [{'parts': [{'text': prompt}]}],
            'generationConfig': generation_config,
        }
        resp = requests.post(url, json=payload, timeout=timeout_sec)
        if resp.status_code != 200:
            err_body = ''
            try:
                err_body = resp.text[:300]
            except Exception:
                pass
            return {
                'status_code': resp.status_code,
                'parsed': None,
                'text': '',
                'error': err_body,
                'model_used': model_to_use,
            }
        body = resp.json()
        text = _extract_text_from_response(body)
        parsed = None
        try:
            parsed = json.loads(text)
        except Exception:
            pass
        return {'status_code': 200, 'parsed': parsed, 'text': text, 'model_used': model_to_use}

    def _call_vertex_api(prompt):
        try:
            from google.oauth2.service_account import Credentials
            from google.auth.transport.requests import Request
        except Exception:
            return None

        project_id = os.getenv('VERTEX_PROJECT_ID', '').strip()
        location = os.getenv('VERTEX_LOCATION', 'us-central1').strip()
        vertex_model = os.getenv('VERTEX_MODEL', 'gemini-2.5-flash').strip()
        cred_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS', '').strip() or 'credenciales_drive.json'
        if not project_id:
            return None

        try:
            creds = Credentials.from_service_account_file(
                cred_path,
                scopes=['https://www.googleapis.com/auth/cloud-platform']
            )
            creds.refresh(Request())
            token = creds.token
            if not token:
                return None

            url = (
                f'https://{location}-aiplatform.googleapis.com/v1/projects/{project_id}/locations/{location}/'
                f'publishers/google/models/{vertex_model}:generateContent'
            )
            payload = {
                'contents': [{'role': 'user', 'parts': [{'text': prompt}]}],
                'generationConfig': {'temperature': 0.0, 'maxOutputTokens': 2048}
            }
            headers = {'Authorization': f'Bearer {token}'}
            resp = requests.post(url, json=payload, headers=headers, timeout=timeout_sec)
            return resp
        except Exception as e:
            print(f'[clasificar_ia] Vertex fallback a API key por error: {e}')
            return None

    def _extract_text_from_response(response_json):
        try:
            cand = response_json.get('candidates', [])
            if not cand:
                return ''
            parts = cand[0].get('content', {}).get('parts', [])
            txt = ''
            for p in parts:
                # gemini-2.5-flash con thinking mode incluye partes thought=True
                # antes del JSON real — las saltamos para no corromper el parse.
                if p.get('thought', False):
                    continue
                txt += str(p.get('text', ''))
            return txt
        except Exception:
            return ''

    def _get_cache_sheet_meta(svc, spreadsheet_id, cache_name):
        meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sh in meta.get('sheets', []):
            props = sh.get('properties', {})
            if props.get('title') == cache_name:
                return props.get('sheetId'), int(props.get('gridProperties', {}).get('rowCount', 0))
        return None, 0

    def _ensure_cache_sheet(svc, spreadsheet_id, cache_name):
        cache_sheet_id_local, _ = _get_cache_sheet_meta(svc, spreadsheet_id, cache_name)
        if cache_sheet_id_local is None:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': [{'addSheet': {'properties': {'title': cache_name}}}]}
            ).execute()

        # Encabezado requerido por negocio.
        # Columna H 'Verificado' (Ola 2c): se marca manualmente con TRUE/SI/X
        # cuando alguien confirma que la fila es correcta. Esos hashes se usan
        # como ejemplos few-shot en próximas corridas.
        # Columna I 'Razon' (Ola 2b): justificación corta que devuelve la IA
        # para la categoría elegida. Sirve de audit trail y ayuda a detectar
        # fallas de criterio sin reabrir el caso.
        # Columna J 'Modelo' (Ola 3): nombre del modelo que produjo la
        # clasificación. Al leer filtramos por modelo actual o por Verificado,
        # de modo que las entradas de modelos viejos quedan invalidadas
        # automáticamente cuando subimos de versión.
        svc.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'{cache_name}!A1:J1',
            valueInputOption='RAW',
            body={'values': [[
                'Id',
                'Sku',
                'Cat original',
                'Cat IA',
                'Fecha',
                'Courier',
                'País',
                'Verificado',
                'Razon',
                'Modelo',
            ]]}
        ).execute()

    def _ensure_rows_capacity_for_append(svc, spreadsheet_id, cache_name, rows_to_add):
        if rows_to_add <= 0:
            return
        cache_sheet_id_local, row_count = _get_cache_sheet_meta(svc, spreadsheet_id, cache_name)
        if cache_sheet_id_local is None:
            return

        # Última fila usada en columna A (incluye encabezado)
        resp = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{cache_name}!A:A'
        ).execute()
        used_rows = len(resp.get('values', []))
        needed_rows = used_rows + rows_to_add + 5
        if needed_rows <= row_count:
            return

        extra_rows = max(500, needed_rows - row_count)
        # Si el libro ya está al tope (10M celdas), batchUpdate falla. No
        # convertimos eso en un crash silencioso del cache: lo logueamos y
        # dejamos que el append posterior falle con un mensaje claro
        # (manejado en _append_cache_sheet).
        try:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    'requests': [{
                        'appendDimension': {
                            'sheetId': cache_sheet_id_local,
                            'dimension': 'ROWS',
                            'length': extra_rows,
                        }
                    }]
                }
            ).execute()
            print(f'[clasificar_ia] Cache IA expandida en +{extra_rows} filas.')
        except Exception as e:
            err_txt = str(e)
            if '10000000 cells' in err_txt or 'limit of 10000000' in err_txt:
                msg = (
                    f"⚠️ No se pudieron reservar +{extra_rows} filas en {spreadsheet_id}: "
                    f"el libro alcanzó el límite de 10M celdas. "
                    f"Mové el cache a un Sheet nuevo (env TAXES_CACHE_SHEET_ID)."
                )
            else:
                msg = f"⚠️ No se pudo expandir cache ({spreadsheet_id}): {err_txt[:200]}"
            print(f'[clasificar_ia] {msg}')
            try:
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
            except Exception:
                pass

    def _leer_cache_sheet(svc, spreadsheet_id, cache_name):
        """Lee el cache aplicando los filtros de versionado.

        Prioridad de aceptación de cada fila:
          1. Si Verificado está marcado (col H) -> aceptamos siempre,
             independientemente del modelo. La verificación humana es el
             único voto persistente de correctitud.
          2. Si Modelo (col J) coincide con MODEL_NAME actual -> aceptamos.
          3. En cualquier otro caso (modelo distinto o vacío) -> ignoramos.

        Modos de override por env var:
          - TAXES_CACHE_REQUIRE_VERIFIED=1 -> sólo entradas verificadas (1).
          - TAXES_CACHE_INCLUDE_OTHER_MODELS=1 -> aceptar también modelos
            distintos del actual (útil para depurar; no usar en prod).
        """
        out = {}
        if not svc or not spreadsheet_id:
            return out
        require_verified = os.getenv('TAXES_CACHE_REQUIRE_VERIFIED', '0').strip().lower() in ('1', 'true', 'yes', 'si')
        include_other_models = os.getenv('TAXES_CACHE_INCLUDE_OTHER_MODELS', '0').strip().lower() in ('1', 'true', 'yes', 'si')
        verified_values = {'TRUE', 'VERDADERO', 'SI', 'SÍ', 'YES', '1', 'X', '✓', 'OK'}
        try:
            _ensure_cache_sheet(svc, spreadsheet_id, cache_name)
            resp = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f'{cache_name}!A2:J'
            ).execute()
            descartadas_modelo = 0
            descartadas_no_verificadas = 0
            aceptadas_verificadas = 0
            aceptadas_modelo = 0
            for row in resp.get('values', []):
                if not row or not row[0].strip():
                    continue
                hash_key = row[0].strip()
                # Cat IA (col D, índice 3)
                categoria_asignada = ''
                if len(row) >= 4 and row[3].strip():
                    categoria_asignada = row[3].strip()
                elif len(row) >= 3 and row[2].strip():
                    categoria_asignada = row[2].strip()
                elif len(row) >= 2 and row[1].strip():
                    categoria_asignada = row[1].strip()
                if not categoria_asignada:
                    continue

                verificado = str(row[7]).strip().upper() if len(row) >= 8 else ''
                modelo_row = str(row[9]).strip() if len(row) >= 10 else ''
                es_verificada = verificado in verified_values

                if require_verified and not es_verificada:
                    descartadas_no_verificadas += 1
                    continue

                if es_verificada:
                    out[hash_key] = _normalizar_categoria(categoria_asignada)
                    aceptadas_verificadas += 1
                    continue

                if not include_other_models:
                    if not modelo_row or modelo_row != MODEL_NAME:
                        descartadas_modelo += 1
                        continue

                out[hash_key] = _normalizar_categoria(categoria_asignada)
                aceptadas_modelo += 1

            print(
                f'[clasificar_ia] Cache Sheet cargado: {len(out)} claves '
                f'(verificadas={aceptadas_verificadas}, modelo={aceptadas_modelo}, '
                f'descartadas_modelo_distinto={descartadas_modelo}, descartadas_no_verificadas={descartadas_no_verificadas})'
            )
        except Exception as e:
            print(f'[clasificar_ia] No se pudo leer cache Sheet: {e}')
        return out

    def _prune_cache_dormant(svc, spreadsheet_id, cache_name):
        """Borra filas del CACHE IA cuyo Modelo (col J) no coincide con el
        actual y que NO están marcadas como Verificadas (col H). Sirve para
        invalidar entradas viejas cuando se sube de modelo y para liberar
        celdas en el libro central.

        Importante: NO borramos filas con razon vacía o con tag
        '[fallback-heuristico]' si su modelo coincide con el actual. Esas
        son resultados del post-procesador que sirven igual y queremos
        reusarlas para no pegarle a la IA en cada corrida. Si la API key
        de Gemini se desbloquea, podemos reprocesar esas explícitamente
        con TAXES_REPROCESS_HEURISTIC=1 (ver código abajo).
        """
        if not svc or not spreadsheet_id:
            return 0
        try:
            _ensure_cache_sheet(svc, spreadsheet_id, cache_name)

            # Resolver sheetId numérico de la pestaña (lo necesitamos para
            # deleteDimension)
            meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            cache_tab_id = None
            for sh in meta.get('sheets', []):
                props = sh.get('properties', {})
                if props.get('title') == cache_name:
                    cache_tab_id = props.get('sheetId')
                    break
            if cache_tab_id is None:
                print(f'[clasificar_ia] Pestaña {cache_name} no encontrada en {spreadsheet_id}, sin prune.')
                return 0

            resp = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f'{cache_name}!A2:J'
            ).execute()
            valores = resp.get('values', [])
            if not valores:
                return 0

            verified_values = {'TRUE', 'VERDADERO', 'SI', 'SÍ', 'YES', '1', 'X', '✓', 'OK'}
            # Permitir reprocesamiento explícito de fallbacks heurísticos
            # cuando la IA volvió a estar disponible.
            reprocess_heuristic = os.getenv('TAXES_REPROCESS_HEURISTIC', '0').strip().lower() in ('1', 'true', 'yes', 'si')
            heuristic_marker = '[fallback-heuristico]'
            indices_a_borrar = []  # 0-indexed (sin contar fila de encabezado)
            for i, row in enumerate(valores):
                # Empty row → considerarla candidata a prune también para evitar gaps
                hash_key = row[0].strip() if row and len(row) > 0 else ''
                modelo_row = str(row[9]).strip() if len(row) >= 10 else ''
                verificado = str(row[7]).strip().upper() if len(row) >= 8 else ''
                razon_row = str(row[8]).strip() if len(row) >= 9 else ''
                if not hash_key:
                    continue
                if verificado in verified_values:
                    continue
                # (a) modelo distinto / sin tag → legacy, fuera
                if not modelo_row or modelo_row != MODEL_NAME:
                    indices_a_borrar.append(i + 1)
                    continue
                # (b) opt-in: borrar fallbacks heurísticos si el usuario quiere
                # reprocesar contra la IA (típicamente tras desbloquear la key).
                if reprocess_heuristic and razon_row.startswith(heuristic_marker):
                    indices_a_borrar.append(i + 1)
                    continue

            if not indices_a_borrar:
                print(f'[clasificar_ia] Prune cache dormant: nada para borrar en {spreadsheet_id}')
                return 0

            # Convertir lista de índices en rangos contiguos para minimizar
            # cantidad de requests de delete y respetar el desplazamiento de
            # índices al borrar (los procesamos de mayor a menor).
            indices_sorted = sorted(set(indices_a_borrar), reverse=True)
            rangos = []  # (start_index, end_index_exclusive) en zero-based del grid
            cur_end = indices_sorted[0]
            cur_start = indices_sorted[0]
            for idx in indices_sorted[1:]:
                if idx == cur_start - 1:
                    cur_start = idx
                else:
                    rangos.append((cur_start, cur_end + 1))
                    cur_start = idx
                    cur_end = idx
            rangos.append((cur_start, cur_end + 1))

            # Cada delete reduce los índices de filas posteriores. Como
            # vamos de mayor a menor en `rangos`, los índices superiores
            # no se ven afectados al ejecutar los menores que aún faltan.
            requests = []
            for start, end_excl in rangos:
                requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': cache_tab_id,
                            'dimension': 'ROWS',
                            'startIndex': start,
                            'endIndex': end_excl,
                        }
                    }
                })

            # Enviar en chunks razonables para no exceder payload
            chunk = 100
            for i in range(0, len(requests), chunk):
                svc.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests[i:i+chunk]}
                ).execute()

            total = len(indices_a_borrar)
            msg = (
                f"🧹 Prune cache dormant: {total} filas eliminadas de {cache_name} "
                f"(modelo != {MODEL_NAME}, o sin razón → fallbacks heurísticos) "
                f"en {spreadsheet_id}"
            )
            print(f'[clasificar_ia] {msg}')
            try:
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
            except Exception:
                pass
            return total
        except Exception as e:
            print(f'[clasificar_ia] No se pudo ejecutar prune cache dormant: {e}')
            try:
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', f'⚠️ No se pudo ejecutar prune cache dormant: {str(e)[:200]}\n')
            except Exception:
                pass
            return 0

    def _leer_ejemplos_verificados(svc, spreadsheet_id, cache_name, pais_actual, courier_actual, max_examples=12):
        """Lee del cache las filas marcadas como Verificado y las devuelve como
        ejemplos few-shot. Filtra por país (y opcionalmente courier) y diversifica
        por categoría para que el prompt no se sesgue a una sola.
        """
        out = []
        if not svc or not spreadsheet_id:
            return out
        try:
            _ensure_cache_sheet(svc, spreadsheet_id, cache_name)
            resp = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f'{cache_name}!A2:H'
            ).execute()
            valores_validos = {'TRUE', 'VERDADERO', 'SI', 'SÍ', 'YES', '1', 'X', '✓', 'OK'}
            pais_target = str(pais_actual or '').strip().lower()
            courier_target = str(courier_actual or '').strip().upper()
            candidatos = []
            for row in resp.get('values', []):
                if len(row) < 8:
                    continue
                verificado = str(row[7]).strip().upper()
                if verificado not in valores_validos:
                    continue
                cat_original = str(row[2]).strip() if len(row) > 2 else ''
                cat_ia = str(row[3]).strip() if len(row) > 3 else ''
                courier_row = str(row[5]).strip().upper() if len(row) > 5 else ''
                pais_row = str(row[6]).strip().lower() if len(row) > 6 else ''
                sku = str(row[1]).strip() if len(row) > 1 else ''
                if not cat_ia:
                    continue
                # Preferimos mismo país; si está vacío en la fila lo aceptamos
                if pais_target and pais_row and pais_row != pais_target:
                    continue
                cat_norm = _normalizar_categoria(cat_ia)
                # Score de prioridad: bonus si coincide courier
                score = 1
                if courier_row and courier_target and courier_row == courier_target:
                    score += 1
                candidatos.append({
                    'cat_original': cat_original,
                    'sku': sku,
                    'categoria': cat_norm,
                    'score': score,
                })

            # Diversificar por categoría: máximo 1 ejemplo por categoría hasta llenar
            candidatos.sort(key=lambda x: x['score'], reverse=True)
            seen_cats = set()
            for c in candidatos:
                if c['categoria'] in seen_cats:
                    continue
                seen_cats.add(c['categoria'])
                out.append(c)
                if len(out) >= max_examples:
                    break
            print(f"[clasificar_ia] Ejemplos verificados cargados: {len(out)} (pais={pais_actual}, courier={courier_actual})")
        except Exception as e:
            print(f'[clasificar_ia] No se pudieron leer ejemplos verificados: {e}')
        return out

    def _append_cache_sheet(svc, spreadsheet_id, cache_name, rows_to_append):
        if not rows_to_append or not svc or not spreadsheet_id:
            return 0
        try:
            _ensure_cache_sheet(svc, spreadsheet_id, cache_name)
            _ensure_rows_capacity_for_append(svc, spreadsheet_id, cache_name, len(rows_to_append))
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            # Columnas: A Id | B Sku | C Cat original | D Cat IA | E Fecha |
            # F Courier | G País | H Verificado (manual) | I Razon (de la IA) |
            # J Modelo (versionado, Ola 3).
            # H se deja vacío para que el operador lo marque manualmente.
            values = [
                [
                    r['hash_key'],
                    r['sku'],
                    r['categoria_original'],
                    r['categoria_asignada'],
                    now,
                    courier,
                    pais,
                    '',
                    str(r.get('razon', '') or '')[:300],
                    MODEL_NAME,
                ]
                for r in rows_to_append
            ]
            svc.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f'{cache_name}!A:J',
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body={'values': values}
            ).execute()
            print(f'[clasificar_ia] Cache Sheet actualizado: +{len(values)} filas (modelo={MODEL_NAME}, sheet={spreadsheet_id})')
            return len(values)
        except Exception as e:
            err_txt = str(e)
            es_cell_limit = '10000000 cells' in err_txt or 'limit of 10000000' in err_txt
            es_permiso = 'does not have permission' in err_txt.lower() or '403' in err_txt
            if es_cell_limit:
                msg = (
                    f"⚠️ No se pudo escribir cache en spreadsheet {spreadsheet_id}: "
                    f"el libro alcanzó el límite de 10M celdas. "
                    f"Solucion: crear un Google Sheet nuevo y setear el env var "
                    f"TAXES_CACHE_SHEET_ID con su ID. Detalle: {err_txt[:200]}"
                )
            elif es_permiso:
                msg = (
                    f"⚠️ No se pudo escribir cache en {spreadsheet_id}: sin permisos. "
                    f"Compartí el sheet con la service account de credenciales_drive.json (Editor). "
                    f"Detalle: {err_txt[:200]}"
                )
            else:
                msg = f"⚠️ No se pudo escribir cache Sheet ({spreadsheet_id}): {err_txt[:300]}"
            print(f'[clasificar_ia] {msg}')
            try:
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
            except Exception:
                pass
            return 0

    def _col_to_a1(col_num):
        result = ''
        while col_num:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _get_sheet_meta(svc, spreadsheet_id, sheet_name):
        meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sh in meta.get('sheets', []):
            props = sh.get('properties', {})
            if props.get('title') == sheet_name:
                grid_props = props.get('gridProperties', {})
                return props.get('sheetId'), grid_props.get('rowCount', 0)
        return None, 0

    def _ensure_sheet_rows_for_append(svc, spreadsheet_id, sheet_name, rows_to_add):
        if rows_to_add <= 0:
            return
        sheet_id_local, row_count = _get_sheet_meta(svc, spreadsheet_id, sheet_name)
        if sheet_id_local is None:
            return

        resp = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:Z'
        ).execute()
        rows = resp.get('values', [])
        used_rows = 0
        for idx in range(len(rows) - 1, -1, -1):
            if any(str(cell).strip() for cell in rows[idx]):
                used_rows = idx + 1
                break

        needed_rows = used_rows + rows_to_add + 5
        if needed_rows <= row_count:
            return

        extra_rows = max(500, needed_rows - row_count)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                'requests': [{
                    'appendDimension': {
                        'sheetId': sheet_id_local,
                        'dimension': 'ROWS',
                        'length': extra_rows,
                    }
                }]
            }
        ).execute()
        print(f'[taxes_thread] Hoja {sheet_name} expandida en +{extra_rows} filas.')

    def _consolidar_input_tm_a_historial(svc, source_spreadsheet_id, source_sheet_name, courier_name, log_widget=None):
        destination_sheet_map = {
            'Courier1': 'ARG',
            'Courier3': 'PE',
            'Courier4': 'ECU',
            'Courier2': 'CR',
            'Courier5': 'UY',
        }
        courier_key = str(courier_name or '').strip().upper()
        destination_sheet_name = destination_sheet_map.get(courier_key)
        if not destination_sheet_name:
            print(f'[taxes_thread] Courier sin consolidado configurado: {courier_name}')
            return 0

        # Spreadsheet del consolidado: configurable por env var. El default
        # mantiene compatibilidad con la instalación actual. Cuando el libro
        # llega al límite de 10M celdas, se reintenta en TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID
        # si está definido.
        primary_sheet_id = (
            os.getenv('TAXES_CONSOLIDADO_SHEET_ID', '').strip()
            or '14lnY-azLXdPIX4BEVarNN7chFu6SmHbSECrB6stkCzc'
        )
        overflow_sheet_id = os.getenv('TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID', '').strip() or ''

        # Leer fuente una sola vez
        try:
            source_resp = svc.spreadsheets().values().get(
                spreadsheetId=source_spreadsheet_id,
                range=f'{source_sheet_name}!A:Z'
            ).execute()
            source_rows = source_resp.get('values', [])
        except Exception as e:
            msg = f'⚠️ No se pudieron leer filas de {source_sheet_name}: {str(e)[:300]}'
            print(f'[taxes_thread] {msg}')
            if log_widget:
                log_widget.insert('end', msg + '\n')
            return 0

        if not source_rows:
            print(f'[taxes_thread] No hay filas para consolidar desde {source_sheet_name}.')
            return 0

        def _ancho_real(rows):
            ancho = 0
            for r in rows:
                if r:
                    # último índice con contenido no vacío
                    for i in range(len(r) - 1, -1, -1):
                        if str(r[i]).strip():
                            ancho = max(ancho, i + 1)
                            break
            return max(ancho, 1)

        def _intentar_consolidar(target_sheet_id):
            """Intenta escribir en el spreadsheet target. Devuelve la cantidad
            de filas escritas o 0 si falló. Si la pestaña destino no existe,
            la crea. Usa values().append() para no tener que pre-asignar
            filas vía batchUpdate (que es lo que dispara el error de 10M).
            """
            # Asegurar que la pestaña destino exista en el target
            try:
                meta = svc.spreadsheets().get(spreadsheetId=target_sheet_id).execute()
                tabs_existentes = {sh.get('properties', {}).get('title') for sh in meta.get('sheets', [])}
                if destination_sheet_name not in tabs_existentes:
                    svc.spreadsheets().batchUpdate(
                        spreadsheetId=target_sheet_id,
                        body={'requests': [{'addSheet': {'properties': {'title': destination_sheet_name}}}]}
                    ).execute()
                    print(f'[taxes_thread] Pestaña {destination_sheet_name} creada en {target_sheet_id}.')
            except Exception as e_meta:
                # Si no podemos consultar/crear el tab, informamos y devolvemos 0
                print(f'[taxes_thread] No se pudo verificar/crear pestaña {destination_sheet_name} en {target_sheet_id}: {e_meta}')
                return 0, str(e_meta)

            # Detectar si la pestaña ya tiene encabezado/datos
            try:
                destination_resp = svc.spreadsheets().values().get(
                    spreadsheetId=target_sheet_id,
                    range=f'{destination_sheet_name}!A:Z'
                ).execute()
                destination_rows = destination_resp.get('values', [])
            except Exception as e_read:
                destination_rows = []
                print(f'[taxes_thread] Aviso leyendo destino {destination_sheet_name} en {target_sheet_id}: {e_read}')

            destination_used_rows = 0
            for idx in range(len(destination_rows) - 1, -1, -1):
                if any(str(cell).strip() for cell in destination_rows[idx]):
                    destination_used_rows = idx + 1
                    break

            # Si la pestaña destino ya tiene encabezado, saltamos la primera fila de la fuente
            if destination_used_rows == 0:
                rows_to_copy = source_rows
            else:
                rows_to_copy = source_rows[1:]

            rows_to_copy = [row for row in rows_to_copy if any(str(cell).strip() for cell in row)]
            if not rows_to_copy:
                print(f'[taxes_thread] No hay datos útiles para consolidar en {destination_sheet_name}.')
                return 0, ''

            # Recortar al ancho real para no inflar celdas vacías al consumo total
            ancho = _ancho_real(rows_to_copy)
            rows_to_copy = [list(r[:ancho]) + [''] * max(0, ancho - len(r)) for r in rows_to_copy]
            end_col = _col_to_a1(ancho)

            try:
                # values().append() extiende la grilla en forma diferida y, en
                # libros sin presión de celdas, no requiere batchUpdate previo.
                svc.spreadsheets().values().append(
                    spreadsheetId=target_sheet_id,
                    range=f'{destination_sheet_name}!A:{end_col}',
                    valueInputOption='USER_ENTERED',
                    insertDataOption='INSERT_ROWS',
                    body={'values': rows_to_copy}
                ).execute()
                return len(rows_to_copy), ''
            except Exception as e_append:
                return 0, str(e_append)

        # Intento primario
        written, err_primary = _intentar_consolidar(primary_sheet_id)
        target_used = primary_sheet_id

        # Si fue cell-limit en primario y hay overflow configurado, reintento
        es_cell_limit_primary = '10000000 cells' in err_primary or 'limit of 10000000' in err_primary
        if written == 0 and es_cell_limit_primary and overflow_sheet_id:
            msg_fb = (
                f"⚠️ Consolidado primario {primary_sheet_id} lleno (10M celdas). "
                f"Reintentando en overflow {overflow_sheet_id}."
            )
            print(f'[taxes_thread] {msg_fb}')
            if log_widget:
                log_widget.insert('end', msg_fb + '\n')
            written, err_overflow = _intentar_consolidar(overflow_sheet_id)
            if written:
                target_used = overflow_sheet_id

        if written:
            msg = (
                f"✅ Consolidado actualizado en {destination_sheet_name}: "
                f"{written} filas copiadas desde {source_sheet_name} "
                f"(destino={target_used})."
            )
            print(f'[taxes_thread] {msg}')
            if log_widget:
                log_widget.insert('end', msg + '\n')
            return written

        # FalCourier2: armamos un mensaje que indique cómo destrabarlo
        err_txt = err_primary or 'desconocido'
        if es_cell_limit_primary:
            if overflow_sheet_id:
                hint = (
                    f"❌ No se pudo consolidar en {destination_sheet_name}: ambos sheets "
                    f"(primary={primary_sheet_id}, overflow={overflow_sheet_id}) fallaron. "
                    f"Detalle primary: {err_txt[:200]}"
                )
            else:
                hint = (
                    f"❌ No se pudo consolidar en {destination_sheet_name}: el libro "
                    f"{primary_sheet_id} alcanzó el límite de 10M celdas. "
                    f"Solución: crear un Google Sheet nuevo y setear "
                    f"TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID con su ID, o mover el "
                    f"consolidado completo a otro libro y setear TAXES_CONSOLIDADO_SHEET_ID."
                )
        else:
            hint = f'⚠️ No se pudo consolidar en {destination_sheet_name}: {err_txt[:300]}'
        print(f'[taxes_thread] {hint}')
        if log_widget:
            log_widget.insert('end', hint + '\n')
        return 0

    def _guardar_progreso_parcial(clasifs):
        if not sheets_service or not sheet_id or not sheet_name:
            return
        try:
            end_row = start_row + len(clasifs) - 1
            range_batch = f'{sheet_name}!{col}{start_row}:{col}{end_row}'
            values_batch = [[c if c else ''] for c in clasifs]
            sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=range_batch,
                valueInputOption='USER_ENTERED',
                body={'values': values_batch}
            ).execute()
        except Exception as e:
            print(f'[clasificar_ia] Error guardando progreso parcial: {e}')

    # Reanudar desde columna L (sin requerir continuidad)
    clasificaciones = [None] * len(rows)
    if sheets_service and sheet_id and sheet_name:
        try:
            existing_range = f"{sheet_name}!{col}{start_row}:{col}{start_row + len(rows) - 1}"
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=existing_range
            ).execute()
            existing_values = result.get('values', [])
            for i, val in enumerate(existing_values):
                if val and val[0].strip():
                    clasificaciones[i] = _corregir_categoria_por_evidencia(
                        _extraer_datos_producto(rows[i]),
                        _normalizar_categoria(val[0])
                    )
            ya_resueltos = sum(1 for c in clasificaciones if c)
            if ya_resueltos:
                print(f'[clasificar_ia] Reanudación: {ya_resueltos}/{len(rows)} ya clasificados en hoja')
        except Exception as e:
            print(f'[clasificar_ia] No se pudieron cargar clasificaciones existentes: {e}')

    # Deduplicar productos pendientes
    hash_to_indices = {}
    hash_to_payload = {}
    for idx, row in enumerate(rows):
        if clasificaciones[idx]:
            continue
        madre, detalles, nombre, sku, categoria_original = _extraer_datos_producto(row)

        key_hash = _cache_key_hash(madre, detalles, nombre, sku)
        hash_to_indices.setdefault(key_hash, []).append(idx)
        if key_hash not in hash_to_payload:
            hash_to_payload[key_hash] = {
                'id': key_hash,
                'madre': madre,
                'detalles': detalles,
                'nombre': nombre,
                'sku': sku,
                'categoria_original': categoria_original,
                'candidatas': _top_categorias_candidatas(madre, detalles, nombre),
            }

    if not hash_to_payload:
        return [c if c else categoria_fallback for c in clasificaciones]

    print(
        f"[clasificar_ia] Pendientes por fila: {sum(1 for c in clasificaciones if not c)} | "
        f"Pendientes únicos: {len(hash_to_payload)}"
    )

    # Bypass total del cache (Ola 3): útil para auditar la calidad real del
    # modelo nuevo sin la inercia de clasificaciones viejas. Cuando está
    # activado igualmente persistimos las nuevas asignaciones para que la
    # próxima corrida sí las aproveche.
    cache_bypass = os.getenv('TAXES_CACHE_BYPASS', '0').strip().lower() in ('1', 'true', 'yes', 'si')

    # Limpieza automática de filas dormant del CACHE IA en cada arranque.
    # "Dormant" = el modelo de la fila no coincide con MODEL_NAME actual y
    # la fila NO fue marcada como Verificada (col H). Verificadas se respetan
    # siempre. La operación es idempotente: si no hay dormants, es un no-op
    # de ~1 segundo. Esto evita que el usuario tenga que pasar env vars
    # manualmente para limpiar cache cuando subimos de modelo o cuando hay
    # entradas viejas con clasificaciones incorrectas.
    # Se puede desactivar puntualmente con TAXES_CACHE_PRUNE_DORMANT=0.
    cache_prune = os.getenv('TAXES_CACHE_PRUNE_DORMANT', '1').strip().lower() in ('1', 'true', 'yes', 'si')
    if cache_prune:
        _prune_cache_dormant(sheets_service, cache_sheet_id, cache_sheet_name)
        if sheet_id and sheet_id != cache_sheet_id:
            _prune_cache_dormant(sheets_service, sheet_id, cache_sheet_name)

    # Cargar cache persistente desde Google Sheet
    if cache_bypass:
        cache_map = {}
        msg_bypass = "[clasificar_ia] TAXES_CACHE_BYPASS=1 → ignorando cache, todo va a IA"
        print(msg_bypass)
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg_bypass + '\n')
    else:
        cache_map = _leer_cache_sheet(sheets_service, cache_sheet_id, cache_sheet_name)
        if sheet_id and sheet_id != cache_sheet_id:
            cache_map_local = _leer_cache_sheet(sheets_service, sheet_id, cache_sheet_name)
            if cache_map_local:
                # Priorizar cache local en caso de colisiones
                cache_map.update(cache_map_local)
                print(f"[clasificar_ia] Cache local adicional cargado: {len(cache_map_local)} claves")

    # Cargar ejemplos verificados (few-shot dinámico, Ola 2c). Llenamos la lista
    # en sitio para que los closures de los prompt builders ya la vean.
    try:
        ejemplos_central = _leer_ejemplos_verificados(
            sheets_service, cache_sheet_id, cache_sheet_name, pais, courier, max_examples=12
        )
        if ejemplos_central:
            verified_examples.extend(ejemplos_central)
        if sheet_id and sheet_id != cache_sheet_id:
            ejemplos_local = _leer_ejemplos_verificados(
                sheets_service, sheet_id, cache_sheet_name, pais, courier, max_examples=12
            )
            # Evitar duplicar la misma categoría que ya esté en el set
            cats_ya = {e.get('categoria') for e in verified_examples}
            for e in ejemplos_local:
                if e.get('categoria') not in cats_ya:
                    verified_examples.append(e)
                    cats_ya.add(e.get('categoria'))
        if verified_examples:
            print(f"[clasificar_ia] Few-shot activo con {len(verified_examples)} ejemplos verificados")
    except Exception as e:
        print(f'[clasificar_ia] Aviso: no se pudieron cargar ejemplos verificados: {e}')

    # Resolver por cache persistente
    cache_hits = 0
    for h, idxs in hash_to_indices.items():
        if h in cache_map:
            cat = _corregir_categoria_por_evidencia(
                hash_to_payload[h],
                _normalizar_categoria(cache_map[h])
            )
            for i in idxs:
                clasificaciones[i] = cat
                cache_hits += 1

    unique_pending = [hash_to_payload[h] for h in hash_to_payload if h not in cache_map]
    print(
        f"[clasificar_ia] Cache hits (filas): {cache_hits} | "
        f"Únicos a enviar a IA: {len(unique_pending)}"
    )

    if not unique_pending:
        return [c if c else categoria_fallback for c in clasificaciones]

    # Control de rate limit global
    rate_lock = threading.Lock()
    last_call = {'ts': 0.0}

    # Circuit breaker compartido entre workers. Si UN batch detecta un error
    # fatal del API (key bloqueada/expirada/inválida), los otros 78 lotes
    # paralelos no tienen sentido que sigan reintentando: la causa es la
    # misma para todos. Acá guardamos la señal y los workers la chequean
    # al inicio de cada intento para abortar limpiamente.
    api_circuit_breaker = {'tripped': False, 'message': '', 'status': 0}

    def _intentar_lote_con_modelo(items_batch, prompt, model_to_use):
        """Una pasada de retries contra `model_to_use`. Devuelve
        (result_map_o_None, ultimo_error_str, ultimo_status). Si result_map es None,
        agotó retries (o hit una falla no-retriable) y el caller decide.
        """
        last_err = ''
        last_status = 0
        # Errores fatales: no tiene sentido reintentar (credenciales, request inválida)
        non_retriable_statuses = {400, 401, 403, 404}
        # Errores que afectan a TODOS los workers (la causa es global): cuando
        # detectamos uno disparamos el circuit breaker para que los demás
        # batches en paralelo aborten sin gastar más retries.
        global_failure_statuses = {400, 401, 403}

        def _trip_breaker(status, err):
            api_circuit_breaker['tripped'] = True
            api_circuit_breaker['message'] = err or ''
            api_circuit_breaker['status'] = status

        for intento in range(max_retries):
            # Circuit breaker: si otro worker ya detectó un fatal global,
            # abortamos sin gastar tiempo.
            if api_circuit_breaker['tripped']:
                return None, api_circuit_breaker['message'], api_circuit_breaker['status']
            try:
                with rate_lock:
                    dt = time.time() - last_call['ts']
                    if dt < min_pause:
                        time.sleep(min_pause - dt)
                    last_call['ts'] = time.time()

                response = None
                if USE_VERTEX:
                    response = _call_vertex_api(prompt)

                if response is not None:
                    status = response.status_code
                    last_status = status
                    if status == 429:
                        wait_time = min((2 ** intento) * 2, 30)
                        print(f'[clasificar_ia] 429 en lote de {len(items_batch)} ({model_to_use}). Reintento en {wait_time}s')
                        time.sleep(wait_time)
                        continue
                    if status in non_retriable_statuses:
                        try:
                            last_err = response.text[:300]
                        except Exception:
                            last_err = ''
                        if status in global_failure_statuses:
                            _trip_breaker(status, last_err)
                            print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}) → fatal global, disparo circuit breaker. Detalle: {last_err}")
                        else:
                            print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}) → fatal, no reintenta. Detalle: {last_err}")
                        return None, last_err, status
                    if status >= 500:
                        # Capturamos el body real para diagnóstico
                        try:
                            last_err = response.text[:300]
                        except Exception:
                            last_err = ''
                        wait_time = min((2 ** intento), 15)
                        print(f'[clasificar_ia] Error {status} en lote ({model_to_use}). Reintento en {wait_time}s. Detalle: {last_err}')
                        time.sleep(wait_time)
                        continue
                    if status != 200:
                        try:
                            last_err = response.text[:300]
                        except Exception:
                            last_err = ''
                        print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}). Detalle: {last_err}")
                        break

                    body = response.json()
                    text = _extract_text_from_response(body)
                    parsed = _parse_json_array(text)
                else:
                    api_result = _call_gemini_api(prompt, batch_response_schema, model_override=model_to_use)
                    status = api_result.get('status_code', 500)
                    last_status = status

                    if status == 429:
                        wait_time = min((2 ** intento) * 2, 30)
                        print(f'[clasificar_ia] 429 en lote de {len(items_batch)} ({model_to_use}). Reintento en {wait_time}s')
                        time.sleep(wait_time)
                        continue

                    if status in non_retriable_statuses:
                        last_err = str(api_result.get('error', ''))[:300]
                        if status in global_failure_statuses:
                            _trip_breaker(status, last_err)
                            print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}) → fatal global, disparo circuit breaker. Detalle: {last_err}")
                        else:
                            print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}) → fatal, no reintenta. Detalle: {last_err}")
                        return None, last_err, status

                    if status >= 500:
                        last_err = str(api_result.get('error', ''))[:300]
                        wait_time = min((2 ** intento), 15)
                        print(f'[clasificar_ia] Error {status} en lote ({model_to_use}). Reintento en {wait_time}s. Detalle: {last_err}')
                        time.sleep(wait_time)
                        continue

                    if status != 200:
                        last_err = str(api_result.get('error', ''))[:300]
                        print(f"[clasificar_ia] HTTP {status} en lote ({model_to_use}). Detalle: {last_err}")
                        break

                    parsed_obj = api_result.get('parsed')
                    if isinstance(parsed_obj, list):
                        parsed = parsed_obj
                    else:
                        parsed = _parse_json_array(str(api_result.get('text', '') or ''))

                if not parsed and intento < max_retries - 1:
                    wait_time = min((2 ** intento) * 2, 30)
                    print(f'[clasificar_ia] Respuesta sin JSON parseable ({model_to_use}). Reintento en {wait_time}s')
                    time.sleep(wait_time)
                    continue

                result_map = {}
                for item in parsed:
                    try:
                        item_id = str(item.get('id', '')).strip()
                        item_cat_raw = str(item.get('categoria', '')).strip()
                        item_razon = str(item.get('razon', '')).strip()
                        item_cat = _normalizar_categoria(item_cat_raw)
                        if item_id:
                            result_map[item_id] = {
                                'categoria_recibida': item_cat_raw,
                                'categoria_asignada': item_cat,
                                'razon': item_razon,
                                'ia_ok': True,
                            }
                    except Exception:
                        continue

                for it in items_batch:
                    if it['id'] not in result_map:
                        result_map[it['id']] = {
                            'categoria_recibida': '',
                            'categoria_asignada': categoria_fallback,
                            'razon': '',
                            'ia_ok': False,
                        }
                return result_map, last_err, last_status or 200
            except Exception as e:
                last_err = str(e)[:300]
                wait_time = min((2 ** intento), 15)
                print(f'[clasificar_ia] Error en lote {model_to_use} (intento {intento + 1}): {last_err}')
                # Si la excepción ya nos dio fatal global (key bloqueada/expirada/inválida),
                # disparamos el breaker y abortamos. No tiene sentido seguir
                # reintentando: la causa es la misma para todos los lotes.
                err_upper = last_err.upper()
                if any(s in err_upper for s in (
                    'PERMISSION_DENIED', 'API KEY', 'API_KEY_INVALID',
                    'API KEY EXPIRED', 'INVALID_ARGUMENT',
                )):
                    _trip_breaker(403, last_err)
                    return None, last_err, 403
                if intento < max_retries - 1:
                    time.sleep(wait_time)

        return None, last_err, last_status

    def _clasificar_lote(items_batch):
        # Si el circuit breaker ya está disparado, no llamamos a la API.
        if api_circuit_breaker['tripped']:
            return {
                it['id']: {
                    'categoria_recibida': '',
                    'categoria_asignada': categoria_fallback,
                    'razon': '',
                    'ia_ok': False,
                }
                for it in items_batch
            }

        prompt = _build_batch_prompt(items_batch)

        # Pasada 1: modelo principal
        result_map, last_err, last_status = _intentar_lote_con_modelo(items_batch, prompt, MODEL_NAME)
        if result_map is not None:
            return result_map

        # Pasada 2: si el principal cayó persistentemente, probamos con el
        # modelo de fallback una sola vez. NO lo intentamos si el problema
        # es de la API key (mismo problema con cualquier modelo) ni si el
        # circuit breaker se disparó por otro lote.
        if api_circuit_breaker['tripped']:
            return {
                it['id']: {
                    'categoria_recibida': '',
                    'categoria_asignada': categoria_fallback,
                    'razon': '',
                    'ia_ok': False,
                }
                for it in items_batch
            }

        if FALLBACK_MODEL and FALLBACK_MODEL != MODEL_NAME:
            print(
                f"[clasificar_ia] Lote {len(items_batch)} agotó retries con {MODEL_NAME}. "
                f"Probando con modelo de fallback {FALLBACK_MODEL}. "
                f"Último error: {last_err}"
            )
            try:
                if 'log_text' in globals() and log_text:
                    log_text.insert('end',
                        f"⚠️ Reintento con modelo fallback {FALLBACK_MODEL}. Detalle: {last_err[:200]}\n"
                    )
            except Exception:
                pass
            result_map_fb, last_err_fb, _ = _intentar_lote_con_modelo(items_batch, prompt, FALLBACK_MODEL)
            if result_map_fb is not None:
                return result_map_fb
            last_err = last_err_fb or last_err

        # Logueo final del error para que el usuario vea por qué falló
        msg_final = (
            f"❌ Lote de {len(items_batch)} items: la IA falló con {MODEL_NAME}"
            + (f" y {FALLBACK_MODEL}" if FALLBACK_MODEL and FALLBACK_MODEL != MODEL_NAME else "")
            + f". Detalle: {last_err[:200]}"
        )
        print(f'[clasificar_ia] {msg_final}')
        try:
            if 'log_text' in globals() and log_text:
                log_text.insert('end', msg_final + '\n')
        except Exception:
            pass

        # Si llegamos acá agotamos retries: marcamos ia_ok=False para que las
        # filas no se persistan al CACHE IA y se reintenten en la próxima
        # corrida (en vez de envenenar el cache con fallbacks heurísticos).
        return {
            it['id']: {
                'categoria_recibida': '',
                'categoria_asignada': categoria_fallback,
                'razon': '',
                'ia_ok': False,
            }
            for it in items_batch
        }

    # Procesar lotes concurrentes
    batches = [unique_pending[i:i + batch_size] for i in range(0, len(unique_pending), batch_size)]
    print(
        f"[clasificar_ia] Clasificación por lotes: {len(batches)} lotes | "
        f"batch_size={batch_size} | workers={max_workers}"
    )

    nuevos_cache = []
    procesados_unicos = 0
    progreso_filas = sum(1 for c in clasificaciones if c)
    revisar_items = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_batch = {executor.submit(_clasificar_lote, b): b for b in batches}
        for fut in as_completed(future_to_batch):
            batch = future_to_batch[fut]
            try:
                batch_result = fut.result()
            except Exception:
                batch_result = {
                    it['id']: {
                        'categoria_recibida': '',
                        'categoria_asignada': categoria_fallback,
                    }
                    for it in batch
                }

            for it in batch:
                h = it['id']
                row_result = batch_result.get(h, {})
                categoria_recibida = ''
                categoria_asignada = categoria_fallback
                razon = ''
                ia_ok = False
                if isinstance(row_result, dict):
                    categoria_recibida = str(row_result.get('categoria_recibida', '')).strip()
                    categoria_asignada = _normalizar_categoria(row_result.get('categoria_asignada', categoria_fallback))
                    razon = str(row_result.get('razon', '')).strip()
                    ia_ok = bool(row_result.get('ia_ok', False))
                else:
                    categoria_asignada = _normalizar_categoria(row_result)

                if categoria_generica and _canon(categoria_asignada) == _canon(categoria_generica):
                    categoria_sugerida = _resolver_los_demas_por_evidencia(it)
                    if categoria_sugerida:
                        categoria_asignada = categoria_sugerida

                categoria_asignada = _corregir_categoria_por_evidencia(it, categoria_asignada)

                for idx in hash_to_indices.get(h, []):
                    if not clasificaciones[idx]:
                        clasificaciones[idx] = categoria_asignada
                        progreso_filas += 1

                if _debe_revisar_los_demas(it, categoria_asignada):
                    revisar_items.append((h, it, categoria_recibida))

                nuevos_cache.append({
                    'hash_key': h,
                    'sku': it.get('sku', ''),
                    'categoria_original': it.get('categoria_original', ''),
                    'categoria_recibida': categoria_recibida,
                    'categoria_asignada': categoria_asignada,
                    'razon': razon,
                    'ia_ok': ia_ok,
                })

            procesados_unicos += len(batch)
            if procesados_unicos % max(1, save_every) == 0:
                print(
                    f"[clasificar_ia] Progreso: filas {progreso_filas}/{len(rows)} | "
                    f"únicos procesados {procesados_unicos}/{len(unique_pending)}"
                )
                _guardar_progreso_parcial(clasificaciones)

    # Re-clasificación de precisión para casos dudosos en la categoría genérica
    if revisar_items:
        print(f"[clasificar_ia] Revisión de precisión para {min(len(revisar_items), max_review_items)} items.")
    revisados = 0
    for h, it, categoria_recibida in revisar_items[:max_review_items]:
        cat_refinada, cat_raw_refinada, razon_refinada = _reclasificar_item(it)
        cat_refinada = _corregir_categoria_por_evidencia(it, cat_refinada)
        if cat_refinada and (not categoria_generica or _canon(cat_refinada) != _canon(categoria_generica)):
            for idx in hash_to_indices.get(h, []):
                clasificaciones[idx] = cat_refinada
            # actualizar también en cache acumulado. Si la IA single devolvió
            # algo (cat_raw_refinada no vacío), marcamos ia_ok=True.
            for r in nuevos_cache:
                if r.get('hash_key') == h:
                    r['categoria_asignada'] = cat_refinada
                    if cat_raw_refinada:
                        r['categoria_recibida'] = cat_raw_refinada
                        r['ia_ok'] = True
                    if razon_refinada:
                        r['razon'] = razon_refinada
                    break
            revisados += 1
    if revisar_items:
        print(f"[clasificar_ia] Revisión precisa completada. Mejorados: {revisados}")

    # Completar faltantes
    clasificaciones = [c if c else categoria_fallback for c in clasificaciones]

    # Persistir cache en Sheet
    # Determinar filas a persistir en cache (permitir forzar escritura para debugging)
    if FORCE_CACHE_WRITE:
        cache_rows = nuevos_cache
        print(f"[clasificar_ia DEBUG] FORCE_CACHE_WRITE active: persistiendo todos los {len(cache_rows)} registros de cache (incluye 'los demas').")
    else:
        if cache_include_los_demas:
            cache_rows = list(nuevos_cache)
        else:
            cache_rows = [
                r for r in nuevos_cache
                if r.get('categoria_asignada') and _canon(r.get('categoria_asignada')) != _canon(categoria_fallback)
            ]
        # Cache fallbacks heurísticos también: cuando la IA no respondió pero
        # el post-procesador dio un resultado decente, lo guardamos con la
        # razon marcada como '[fallback-heuristico]'. Eso permite reusarlo
        # en próximas corridas (rápido) y a la vez identificarlo por el tag.
        for r in cache_rows:
            if not r.get('ia_ok'):
                if not str(r.get('razon', '') or '').strip():
                    r['razon'] = '[fallback-heuristico]'
        print(f"[clasificar_ia DEBUG] nuevos_cache={len(nuevos_cache)} -> cache_rows_filtradas={len(cache_rows)} (cache_include_los_demas={cache_include_los_demas})")

    written_cache_rows = _append_cache_sheet(sheets_service, cache_sheet_id, cache_sheet_name, cache_rows)
    cache_target = cache_sheet_id
    if written_cache_rows == 0 and cache_rows and sheet_id and sheet_id != cache_sheet_id:
        msg_fb = (
            f"[clasificar_ia] Fallback de escritura cache: central sin escritura, "
            f"intentando en sheet local {sheet_id}."
        )
        print(msg_fb)
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg_fb + '\n')
        written_cache_rows = _append_cache_sheet(sheets_service, sheet_id, cache_sheet_name, cache_rows)
        if written_cache_rows > 0:
            cache_target = sheet_id
    if cache_rows and written_cache_rows == 0:
        msg_fail = (
            f"⚠️ No se persistió ninguna fila en CACHE IA. Revisar permisos/cell-limit "
            f"en {cache_sheet_id} y {sheet_id}."
        )
        print(f'[clasificar_ia] {msg_fail}')
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg_fail + '\n')

    print(
        f"[clasificar_ia] Cache persistida: solicitadas={len(cache_rows)} | escritas={written_cache_rows} | "
        f"destino={cache_target}"
    )
    if 'log_text' in globals() and log_text:
        log_text.insert(
            'end',
            f"[clasificar_ia] Cache: hits={cache_hits}, nuevas={len(cache_rows)}, escritas={written_cache_rows}\n"
        )

    tiempo_fin = time.time()
    duracion_total = tiempo_fin - tiempo_inicio
    minutos, segundos = divmod(int(duracion_total), 60)
    
    filas_procesadas_ia = sum(1 for c in clasificaciones if c and _canon(c) != _canon(categoria_fallback))
    filas_los_demas = sum(1 for c in clasificaciones if _canon(c) == _canon(categoria_fallback))
    
    con_razon = sum(1 for r in nuevos_cache if str(r.get('razon', '') or '').strip())
    msg_final = (
        f"[clasificar_ia] ✅ Clasificación completada: {len(clasificaciones)} filas | "
        f"cache_hits={cache_hits} | únicos_ia={len(unique_pending)} | "
        f"ia_procesadas={filas_procesadas_ia} | los_demas={filas_los_demas} | "
        f"few_shot={len(verified_examples)} | con_razon={con_razon}/{len(nuevos_cache)} | "
        f"tiempo_total={minutos}m {segundos}s"
    )
    print(msg_final)
    if 'log_text' in globals() and log_text:
        log_text.insert('end', f"\n{msg_final}\n")

    return clasificaciones


def _Courier3_parse_excel(fh, fecha_inicio_dt, fecha_fin_dt):
    """
    Parsea el Excel de Courier3 con estructura horizontal/vertical compleja.
    Soporta GUIA en formato 1600XXXXXX (10 dígitos) y PEXXXXXXXXX.
    Tasa PEN/USD tomada de celdas con patrón 'X 3.NNN = S/' en el mismo sheet.
    """
    import re
    import pandas as pd
    from datetime import datetime

    df_raw = pd.read_excel(fh, header=None, dtype=str, sheet_name=0)

    rate_pattern = re.compile(r'[Xx]\s*([\d]{1,2}[.,][\d]{3,})\s*=\s*[Ss]', re.IGNORECASE)
    # DD/MM/YYYY o DD/MM/YY (texto manual)
    date_pattern     = re.compile(r'(\d{1,2})[/](\d{1,2})[/](\d{2,4})')
    # YYYY-MM-DD … (pandas convierte celdas-fecha de Excel a ISO al usar dtype=str)
    date_iso_pattern = re.compile(r'(\d{4})-(\d{2})-(\d{2})')
    # Formatos de GUIA aceptados: 10 dígitos  o  PE + hasta 12 dígitos
    guia_numeric   = re.compile(r'^\d{7,12}$')
    guia_pe_format = re.compile(r'^PE\d{6,12}$', re.IGNORECASE)

    def _parse_date(cell_str):
        """Intenta parsear una fecha en formato DD/MM/YYYY o YYYY-MM-DD. Devuelve datetime o None."""
        dm = date_pattern.search(cell_str)
        if dm:
            try:
                d, mo, y = int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
                if y < 100:
                    y += 2000
                return datetime(y, mo, d)
            except Exception:
                pass
        dm_iso = date_iso_pattern.search(cell_str)
        if dm_iso:
            try:
                y, mo, d = int(dm_iso.group(1)), int(dm_iso.group(2)), int(dm_iso.group(3))
                return datetime(y, mo, d)
            except Exception:
                pass
        return None

    # --- Paso 1: Tasas de cambio ---
    # Escanea todo el sheet; la fecha asociada a cada tasa es la de esa misma fila
    # o, si la tasa está en una celda multi-dato, la más cercana a la izquierda.
    tasas = []
    last_seen_date = None

    for row_idx in range(len(df_raw)):
        row_date = None
        # Primera pasada: encontrar fecha en la fila
        for col_idx in range(len(df_raw.columns)):
            cell = str(df_raw.iloc[row_idx, col_idx]).strip()
            if not cell or cell.lower() == 'nan':
                continue
            if not rate_pattern.search(cell):
                parsed = _parse_date(cell)
                if parsed:
                    row_date = parsed
                    last_seen_date = row_date
                    break

        # Segunda pasada: extraer tasas de la misma fila
        for col_idx in range(len(df_raw.columns)):
            cell = str(df_raw.iloc[row_idx, col_idx]).strip()
            if not cell or cell.lower() == 'nan':
                continue
            m = rate_pattern.search(cell)
            if m:
                try:
                    rate_str = m.group(1).replace(',', '.')
                    rate = float(rate_str)
                    if 3.0 <= rate <= 5.0:
                        # Intentar extraer fecha del mismo token
                        tasa_date = _parse_date(cell) or row_date or last_seen_date
                        if tasa_date:
                            tasas.append((tasa_date, rate))
                except (ValueError, AttributeError):
                    pass

    tasas.sort(key=lambda x: x[0])
    print(f"[Courier3] {len(tasas)} tasas de cambio identificadas")

    def get_rate(target_dt):
        if not tasas:
            return None
        applicable = [t for t in tasas if t[0] <= target_dt]
        return applicable[-1][1] if applicable else tasas[0][1]

    def normalize_guia(raw):
        """Devuelve la GUIA normalizada o None si no es válida."""
        s = str(raw).strip()
        if not s or s.lower() == 'nan':
            return None
        if guia_pe_format.match(s.upper()):
            return s.upper()
        if guia_numeric.match(s):
            return s
        # Último recurso: solo dígitos >= 7 chars
        digits = ''.join(c for c in s if c.isdigit())
        if len(digits) >= 7:
            return digits
        return None

    # --- Paso 2: Encontrar todos los bloques con cabecera GUIA ---
    resultados = []
    processed_guia = set()  # (header_row, guia_col) ya procesados

    for row_idx in range(len(df_raw)):
        for col_idx in range(len(df_raw.columns)):
            cell_val = str(df_raw.iloc[row_idx, col_idx]).strip().upper()
            if cell_val != "GUIA":
                continue
            if (row_idx, col_idx) in processed_guia:
                continue
            processed_guia.add((row_idx, col_idx))

            guia_col   = col_idx
            header_row = row_idx

            # Buscar columna FECHA hasta 8 columnas a la izquierda del GUIA
            fecha_col = None
            for sc in range(max(0, guia_col - 8), guia_col):
                hdr = str(df_raw.iloc[header_row, sc]).strip().upper()
                if "FECHA" in hdr:
                    fecha_col = sc
                    break
            if fecha_col is None:
                continue

            monto_col = fecha_col + 2  # FECHA | Descripción | MONTO | Nº op | GUIA
            if monto_col >= len(df_raw.columns):
                continue

            # Recorrer filas de datos; tolerar hasta 3 filas consecutivas sin fecha
            consecutive_no_date = 0
            MAX_EMPTY = 3

            for data_row in range(header_row + 1, len(df_raw)):
                fecha_val = str(df_raw.iloc[data_row, fecha_col]).strip()
                guia_val  = str(df_raw.iloc[data_row, guia_col]).strip()

                # Ambas vacías → posible fin de bloque
                if (not fecha_val or fecha_val.lower() == 'nan') and \
                   (not guia_val  or guia_val.lower()  == 'nan'):
                    consecutive_no_date += 1
                    if consecutive_no_date >= MAX_EMPTY:
                        break
                    continue

                fecha_dt = _parse_date(fecha_val) if fecha_val and fecha_val.lower() != 'nan' else None
                if not fecha_dt:
                    # Fila de totales / vacía parcial → saltar sin contar como vacía
                    continue
                consecutive_no_date = 0

                # Filtrar por rango de fechas del usuario
                if not (fecha_inicio_dt.date() <= fecha_dt.date() <= fecha_fin_dt.date()):
                    continue

                guia_final = normalize_guia(guia_val)
                if not guia_final:
                    continue

                monto_raw = str(df_raw.iloc[data_row, monto_col]).strip() \
                            if monto_col < len(df_raw.columns) else ""
                try:
                    monto_num = float(monto_raw.replace(',', '').replace(' ', '').replace('$', ''))
                except (ValueError, AttributeError):
                    continue
                if monto_num == 0:
                    continue

                rate = get_rate(fecha_dt)
                if not rate:
                    continue

                usd_amount = round(abs(monto_num) / rate, 2)
                # Formato INPUT Courier: Nro.Guia | vacío | vacío | vacío | TotalUSD | vacío
                resultados.append({0: guia_final, 1: "", 2: "", 3: "", 4: usd_amount, 5: ""})

    if not resultados:
        return pd.DataFrame(columns=[0, 1, 2, 3, 4, 5])

    return pd.DataFrame(resultados)


def _taxes_Courier3_procesar(sharepoint_url, sheet_id, sheet_name,
                          fecha_inicio_dt, fecha_fin_dt,
                          sheets_service, log_text, ordenes_lista):
    """Descarga el archivo Courier3 desde SharePoint, parsea y sube a Google Sheets INPUT Courier."""
    import requests as _req
    import io as _io
    import math
    import decimal

    if log_text:
        log_text.insert('end', "⬇️ Descargando archivo Courier3 desde SharePoint...\n")
    print("[Courier3] Descargando archivo desde SharePoint...")

    session = _req.Session()
    resp = session.get(sharepoint_url, allow_redirects=True, timeout=120)
    if resp.status_code != 200:
        raise ValueError(
            f"No se pudo descargar el archivo Courier3 (HTTP {resp.status_code}). "
            "Verifique que el enlace de SharePoint esté configurado como "
            "'Cualquier persona con el vínculo puede ver'."
        )

    fh = _io.BytesIO(resp.content)
    if log_text:
        log_text.insert('end', "📊 Parseando datos Courier3...\n")
    print("[Courier3] Parseando Excel...")

    df = _Courier3_parse_excel(fh, fecha_inicio_dt, fecha_fin_dt)

    if df.empty:
        msg = "No se encontraron registros Courier3 en el rango de fechas seleccionado."
        print(f"[Courier3] {msg}")
        if log_text:
            log_text.insert('end', msg + '\n')
        return

    print(f"[Courier3] {len(df)} registros encontrados.")
    if log_text:
        log_text.insert('end', f"✅ {len(df)} registros Courier3 encontrados.\n")

    # Limpiar rango en Google Sheets
    clear_range = f"{sheet_name}!A2:F"
    print(f"[Courier3] Limpiando {clear_range}...")
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=clear_range
    ).execute()

    def _conv(obj):
        if isinstance(obj, float) and math.isnan(obj):
            return ""
        if isinstance(obj, decimal.Decimal):
            return float(obj)
        return obj

    values = [[_conv(c) for c in row] for row in df.values.tolist()]
    start_row = 2
    end_row = start_row + len(values) - 1
    range_destino = f"{sheet_name}!A{start_row}:F{end_row}"
    print(f"[Courier3] Subiendo datos a {range_destino}")
    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_destino,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

    msg = f"✅ Datos Courier3 subidos a {sheet_name}: {len(values)} filas."
    print(f"[Courier3] {msg}")
    if log_text:
        log_text.insert('end', msg + '\n')

    # Registrar para el finally
    ordenes_lista.extend([str(row[0]) for row in values if str(row[0]).strip()])


def _process_Courier4_files(drive_service, folder_id, fecha_inicio, fecha_fin, log_text):
    """Procesa archivos de Courier4 en PDF/XLS/XLSX.

    Devuelve una lista de dicts con:
    - identificador: valor para la columna A de INPUT Courier
    - total_usd: valor para la columna E
    - origen: legacy_pdf o tabla
    - archivo: nombre del archivo origen
    """
    import io
    import math
    import datetime
    import pdfplumber
    import pandas as pd
    from googleapiclient.http import MediaIoBaseDownload

    def _canon(v):
        txt = ' '.join(str(v or '').strip().split()).lower()
        txt = unicodedata.normalize('NFD', txt)
        return ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')

    def _parse_amount(raw):
        if raw is None:
            return None
        if isinstance(raw, (int, float)) and not (isinstance(raw, float) and math.isnan(raw)):
            return float(raw)
        txt = str(raw).strip()
        if not txt:
            return None
        txt = txt.replace('$', '').replace('USD', '').replace('usd', '').replace(' ', '')
        txt = txt.replace('\u00a0', '')
        txt = re.sub(r'[^\d,\.\-]', '', txt)
        if not txt:
            return None
        if ',' in txt and '.' in txt:
            if txt.rfind(',') > txt.rfind('.'):
                txt = txt.replace('.', '').replace(',', '.')
            else:
                txt = txt.replace(',', '')
        elif ',' in txt:
            partes = txt.split(',')
            if len(partes[-1]) in (1, 2):
                txt = txt.replace('.', '').replace(',', '.')
            else:
                txt = txt.replace(',', '')
        try:
            return float(txt)
        except Exception:
            return None

    def _normalize_order(raw):
        txt = str(raw or '').strip()
        if not txt or txt.lower() == 'nan':
            return ''
        txt = txt.replace('\n', ' ').replace('\r', ' ')
        txt = ' '.join(txt.split())
        txt = txt.replace('HAWB', '').replace('GUIA', '')
        txt = txt.replace(':', ' ').replace('#', ' ')
        txt = ' '.join(txt.split())
        return txt.upper()

    def _detect_table_rows(rows):
        """
        Detecta columnas HAWB/GUIA y sus correspondientes columnas de precio.
        Prioriza: HAWB (→ Autoliquidación) sobre GUIA (→ Impuesto).
        Busca por coincidencia EXACTA de encabezados (sin substring).
        """
        hawb_idx = None
        guia_idx = None
        impuesto_idx = None
        autoliquidacion_idx = None
        header_row = None

        # Buscar en las primeras filas: encabezados exactos
        for idx, row in enumerate(rows[:12]):
            for col_idx, cell in enumerate(row):
                canon_cell = _canon(cell)
                cell_str = str(cell or '').strip()

                # Búsqueda EXACTA de HAWB
                if hawb_idx is None and canon_cell == 'hawb':
                    hawb_idx = col_idx
                    if header_row is None:
                        header_row = idx

                # Búsqueda EXACTA de GUIA (no "Guía Aérea")
                if guia_idx is None and canon_cell == 'guia':
                    guia_idx = col_idx
                    if header_row is None:
                        header_row = idx

                # Búsqueda EXACTA de Impuesto
                if impuesto_idx is None and canon_cell == 'impuesto':
                    impuesto_idx = col_idx
                    if header_row is None:
                        header_row = idx

                # Búsqueda EXACTA de Autoliquidación
                if autoliquidacion_idx is None and (canon_cell == 'autoliquidacion' or canon_cell == 'autoliquidación'):
                    autoliquidacion_idx = col_idx
                    if header_row is None:
                        header_row = idx

        # Determinar qué combinación tenemos y extraer datos
        registros = []

        # Prioridad 1: HAWB + Autoliquidación
        if hawb_idx is not None and autoliquidacion_idx is not None:
            data_start = header_row + 1 if header_row is not None else 0
            for row in rows[data_start:]:
                if not any(str(c).strip() for c in row):
                    continue
                order_raw = row[hawb_idx] if hawb_idx < len(row) else ''
                amount_raw = row[autoliquidacion_idx] if autoliquidacion_idx < len(row) else ''
                order = _normalize_order(order_raw)
                amount = _parse_amount(amount_raw)
                if not order or amount is None:
                    continue
                registros.append((order, amount))

        # Prioridad 2: GUIA + Impuesto (si no encontramos HAWB+Autoliquidación)
        elif guia_idx is not None and impuesto_idx is not None:
            data_start = header_row + 1 if header_row is not None else 0
            for row in rows[data_start:]:
                if not any(str(c).strip() for c in row):
                    continue
                order_raw = row[guia_idx] if guia_idx < len(row) else ''
                amount_raw = row[impuesto_idx] if impuesto_idx < len(row) else ''
                order = _normalize_order(order_raw)
                amount = _parse_amount(amount_raw)
                if not order or amount is None:
                    continue
                registros.append((order, amount))

        # Si encontramos datos válidos, retornar deduplicados
        if registros:
            resumen = {}
            for o, a in registros:
                resumen.setdefault(o, 0)
                try:
                    resumen[o] += float(a)
                except Exception:
                    pass
            return list(resumen.items())

        # Fallback heurístico: buscar patrones de orden y montos en cualquier columna/fila
        registros_fb = []
        order_re = re.compile(r'([A-Za-z]{1,4}\d{6,}|\d{2,}-\d{4,}|\d{6,})', re.IGNORECASE)
        amount_re = re.compile(r'\$|USD|\d+[.,]\d{1,}|\d{1,3}(?:[.,]\d{3})+')

        for r_idx, row in enumerate(rows):
            if not any(str(c).strip() for c in row):
                continue
            for c_idx, cell in enumerate(row):
                s = str(cell or "").strip()
                if not s:
                    continue
                m = order_re.search(s)
                if not m:
                    continue
                order_raw = m.group(0)
                order_val = _normalize_order(order_raw)

                found_amount = None
                for look_c in range(max(0, c_idx - 3), min(len(row), c_idx + 4)):
                    if look_c == c_idx:
                        continue
                    cell2 = row[look_c]
                    if cell2 is None:
                        continue
                    if not amount_re.search(str(cell2)):
                        continue
                    parsed = _parse_amount(cell2)
                    if parsed is not None and parsed > 0:
                        try:
                            if str(parsed).replace('.', '').replace(',', '') == re.sub(r'\D', '', order_raw):
                                continue
                        except Exception:
                            pass
                        found_amount = parsed
                        break
                if found_amount is None:
                    for look_r in range(max(0, r_idx - 2), min(len(rows), r_idx + 3)):
                        if look_r == r_idx:
                            continue
                        row2 = rows[look_r]
                        if not row2:
                            continue
                        for cell2 in row2:
                            if cell2 is None:
                                continue
                            if not amount_re.search(str(cell2)):
                                continue
                            parsed = _parse_amount(cell2)
                            if parsed is not None and parsed > 0:
                                try:
                                    if str(parsed).replace('.', '').replace(',', '') == re.sub(r'\D', '', order_raw):
                                        continue
                                except Exception:
                                    pass
                                found_amount = parsed
                                break
                        if found_amount is not None:
                            break

                if found_amount is not None:
                    registros_fb.append((order_val, found_amount))

        seen = set()
        final = []
        for o, a in registros_fb:
            key = (o, float(a))
            if key in seen:
                continue
            seen.add(key)
            final.append((o, float(a)))

        return final

    def _extract_from_pdf(pdf_bytes, archivo_nombre):
        with pdfplumber.open(pdf_bytes) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    rows = [list(r) for r in table if r]
                    registros = _detect_table_rows(rows)
                    if registros:
                        return [{
                            'identificador': order,
                            'total_usd': amount,
                            'origen': 'tabla',
                            'archivo': archivo_nombre,
                        } for order, amount in registros]

                texto_pagina = page.extract_text() or ''
                if texto_pagina:
                    rows_texto = [
                        re.split(r'\t+|\s{2,}', line.strip())
                        for line in texto_pagina.splitlines()
                        if line.strip()
                    ]
                    registros = _detect_table_rows(rows_texto)
                    if registros:
                        return [{
                            'identificador': order,
                            'total_usd': amount,
                            'origen': 'tabla',
                            'archivo': archivo_nombre,
                        } for order, amount in registros]

            texto = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        vm = re.search(r'VUELO\s+(\d+)', texto, re.IGNORECASE)
        tm = re.search(r'VALOR\s+TOTAL\s+USD\s*:?\s*([\d,]+\.?\d*)', texto, re.IGNORECASE)
        if vm and tm:
            return [{
                'identificador': f'Courier4 {vm.group(1)}',
                'total_usd': float(tm.group(1).replace(',', '')),
                'origen': 'legacy_pdf',
                'archivo': archivo_nombre,
            }]
        return []

    def _extract_from_excel(excel_bytes, archivo_nombre):
        registros = []
        try:
            xls = pd.ExcelFile(excel_bytes)
        except Exception:
            return registros

        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
            except Exception:
                continue

            rows = df.fillna('').astype(str).values.tolist()
            sheet_records = _detect_table_rows(rows)
            if not sheet_records:
                continue

            registros.extend({
                'identificador': order,
                'total_usd': amount,
                'origen': 'tabla',
                'archivo': archivo_nombre,
            } for order, amount in sheet_records)

        return registros

    datos = []
    try:
        q = (
            f"'{folder_id}' in parents and trashed=false and ("
            "mimeType='application/pdf' or "
            "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or "
            "mimeType='application/vnd.ms-excel')"
        )
        results = drive_service.files().list(
            q=q,
            fields="files(id,name,modifiedTime,mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = [
            f for f in results.get('files', [])
            if fecha_inicio <= datetime.datetime.strptime(f['modifiedTime'][:10], "%Y-%m-%d") <= fecha_fin
        ]

        for f in files:
            try:
                req = drive_service.files().get_media(fileId=f['id'])
                fh = io.BytesIO()
                dl = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    _, done = dl.next_chunk()
                fh.seek(0)

                mime_type = f.get('mimeType', '')
                nombre = f.get('name', 'archivo')
                es_excel = mime_type in (
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'application/vnd.ms-excel'
                ) or nombre.lower().endswith(('.xls', '.xlsx'))

                if es_excel:
                    registros_file = _extract_from_excel(fh, nombre)
                else:
                    registros_file = _extract_from_pdf(fh, nombre)

                if registros_file:
                    datos.extend(registros_file)
                    log_text.insert('end', f"✅ {nombre}: {len(registros_file)} registro(s) detectado(s)\n")
                else:
                    log_text.insert('end', f"⚠️ {nombre}: no se detectaron columnas HAWB/GUIA + Impuesto/Autoliquidación o formato legacy\n")
            except Exception as e:
                log_text.insert('end', f"⚠️ {f.get('name', 'archivo')}: error leyendo archivo ({e})\n")
    except Exception as e:
        log_text.insert('end', f"⚠️ Error listando archivos Courier4: {e}\n")

    if datos:
        resumen_ordenes = {}
        origen_por_identificador = {}
        for item in datos:
            key = item['identificador']
            resumen_ordenes.setdefault(key, 0.0)
            resumen_ordenes[key] += float(item['total_usd'])
            origen_por_identificador.setdefault(key, item.get('origen', 'tabla'))

        datos = [
            {
                'identificador': identificador,
                'total_usd': total,
                'origen': origen_por_identificador.get(identificador, 'tabla'),
            }
            for identificador, total in sorted(resumen_ordenes.items())
        ]

        # Filtrar solo órdenes con formato EC######
        datos_validos = [
            d for d in datos
            if re.match(r'^EC\d{6,}$', str(d['identificador']).strip(), re.IGNORECASE)
        ]
        datos_invalidos = [d for d in datos if d not in datos_validos]

        if datos_invalidos:
            log_text.insert('end', f"⚠️ {len(datos_invalidos)} identificador(es) rechazado(s) por formato inválido (esperado: EC######)\n")

        log_text.insert('end', f"📊 Resultado: {len(datos_validos)} identificador(es) válido(s)\n")

        datos = datos_validos

    return datos


def _process_Courier4_pdfs(drive_service, folder_id, fecha_inicio, fecha_fin, log_text):
    """Compatibilidad hacia atrás: usa el parser nuevo para PDF/XLS/XLSX."""
    return _process_Courier4_files(drive_service, folder_id, fecha_inicio, fecha_fin, log_text)


def _taxes_col_to_a1(col_num):
    result = ''
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _taxes_get_sheet_meta(svc, spreadsheet_id, sheet_name):
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sh in meta.get('sheets', []):
        props = sh.get('properties', {})
        if props.get('title') == sheet_name:
            grid_props = props.get('gridProperties', {})
            return props.get('sheetId'), int(grid_props.get('rowCount', 0))
    return None, 0


def _taxes_ensure_sheet_rows_for_append(svc, spreadsheet_id, sheet_name, rows_to_add):
    if rows_to_add <= 0:
        return

    sheet_id_local, row_count = _taxes_get_sheet_meta(svc, spreadsheet_id, sheet_name)
    if sheet_id_local is None:
        return

    resp = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f'{sheet_name}!A:Z'
    ).execute()
    rows = resp.get('values', [])

    used_rows = 0
    for idx in range(len(rows) - 1, -1, -1):
        if any(str(cell).strip() for cell in rows[idx]):
            used_rows = idx + 1
            break

    needed_rows = used_rows + rows_to_add + 5
    if needed_rows <= row_count:
        return

    extra_rows = max(500, needed_rows - row_count)
    svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'requests': [{
                'appendDimension': {
                    'sheetId': sheet_id_local,
                    'dimension': 'ROWS',
                    'length': extra_rows,
                }
            }]
        }
    ).execute()
    print(f'[taxes_thread] Hoja {sheet_name} expandida en +{extra_rows} filas.')


def _taxes_consolidar_input_tm_a_historial(svc, source_spreadsheet_id, source_sheet_name, courier_name,
                                            log_widget=None, key_cols=(0,), tipo=None,
                                            destination_sheet_name_override=None):
    """Consolida una hoja fuente al historial centralizado.

    key_cols: tupla de índices de columna que forman la clave compuesta para deduplicación.
              (0,) para solo col A (Guide), (0, 1) para col A + col B (orden + SKU).
    tipo: si se indica ('Pre' o 'Pos'), se agrega como columna extra al final de cada fila.
    destination_sheet_name_override: nombre de pestaña destino explícito, omite el mapeo por courier.
    """
    destination_sheet_map = {
        'Courier1': 'ARG',
        'Courier3': 'PE',
        'Courier4': 'ECU',
        'Courier2': 'CR',
        'Courier5': 'UY',
    }

    courier_key = str(courier_name or '').strip().upper()
    destination_sheet_name = destination_sheet_name_override or destination_sheet_map.get(courier_key)
    if not destination_sheet_name:
        print(f'[taxes_thread] Courier sin consolidado configurado: {courier_name}')
        return 0

    # Cuando tipo está activo, se incluye en la clave compuesta para que filas
    # Pre y Pos de la misma orden+SKU nunca colisionen entre sí: un run Pos
    # nunca pisa una fila Pre y viceversa.
    # Clave para filas FUENTE (sin tipo todavía, key_cols posiciones originales).
    def _make_src_key(row):
        parts = [str(row[c]).strip().upper() if len(row) > c else '' for c in key_cols]
        if tipo:
            parts = [tipo.strip().upper()] + parts
        return '|'.join(parts)

    # Clave para filas DESTINO: col A = tipo (si activo), cols B+ = datos.
    dest_key_cols = tuple(c + 1 for c in key_cols) if tipo else key_cols

    def _make_dest_key(row):
        parts = [str(row[c]).strip().upper() if len(row) > c else '' for c in dest_key_cols]
        if tipo:
            tipo_dest = str(row[0]).strip().upper() if row else ''
            parts = [tipo_dest] + parts
        return '|'.join(parts)

    # Spreadsheet del consolidado: configurable por env var (Ola 3).
    # Default mantiene compatibilidad con la instalación actual. Cuando el libro
    # llega al límite de 10M celdas y hay overflow definido, las filas NUEVAS
    # se redirigen automáticamente a TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID.
    consolidado_sheet_id = (
        os.getenv('TAXES_CONSOLIDADO_SHEET_ID', '').strip()
        or '14lnY-azLXdPIX4BEVarNN7chFu6SmHbSECrB6stkCzc'
    )
    overflow_sheet_id = os.getenv('TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID', '').strip() or ''
    try:
        source_resp = svc.spreadsheets().values().get(
            spreadsheetId=source_spreadsheet_id,
            range=f'{source_sheet_name}!A:Z'
        ).execute()
        source_rows = source_resp.get('values', [])
        if not source_rows:
            print(f'[taxes_thread] No hay filas para consolidar desde {source_sheet_name}.')
            return 0

        destination_resp = svc.spreadsheets().values().get(
            spreadsheetId=consolidado_sheet_id,
            range=f'{destination_sheet_name}!A:Z'
        ).execute()
        destination_rows = destination_resp.get('values', [])

        destination_used_rows = 0
        for idx in range(len(destination_rows) - 1, -1, -1):
            if any(str(cell).strip() for cell in destination_rows[idx]):
                destination_used_rows = idx + 1
                break

        if destination_used_rows == 0:
            rows_to_copy = source_rows
        else:
            rows_to_copy = source_rows[1:]

        rows_to_copy = [row for row in rows_to_copy if any(str(cell).strip() for cell in row)]
        if not rows_to_copy:
            print(f'[taxes_thread] No hay datos útiles para consolidar en {destination_sheet_name}.')
            return 0

        # Deduplicación en destino: usa _make_dest_key (desplazada si tipo activo).
        # Si la misma clave aparece más de una vez, conservar la primera
        # (será actualizada) y marcar las demás para borrar.
        existing_refs_map = {}  # {key: row_number 1-indexed}
        dest_duplicate_indices = []  # índices 0-indexed de filas a eliminar del destino
        for idx, dest_row in enumerate(destination_rows):
            if dest_row and len(dest_row) > 0:
                ref_value = _make_dest_key(dest_row)
                if ref_value:
                    if ref_value in existing_refs_map:
                        dest_duplicate_indices.append(idx)
                    else:
                        existing_refs_map[ref_value] = idx + 1  # fila 1-indexed

        # Eliminar filas duplicadas del destino (de abajo a arriba para no
        # desplazar índices durante el borrado).
        if dest_duplicate_indices:
            dest_sheet_gid, _ = _taxes_get_sheet_meta(svc, consolidado_sheet_id, destination_sheet_name)
            if dest_sheet_gid is not None:
                delete_requests = []
                for row_idx in sorted(dest_duplicate_indices, reverse=True):
                    delete_requests.append({
                        'deleteDimension': {
                            'range': {
                                'sheetId': dest_sheet_gid,
                                'dimension': 'ROWS',
                                'startIndex': row_idx,
                                'endIndex': row_idx + 1,
                            }
                        }
                    })
                try:
                    svc.spreadsheets().batchUpdate(
                        spreadsheetId=consolidado_sheet_id,
                        body={'requests': delete_requests}
                    ).execute()
                    print(f'[taxes_thread] {len(dest_duplicate_indices)} filas duplicadas eliminadas de {destination_sheet_name}.')
                except Exception as e_del:
                    print(f'[taxes_thread] Error eliminando duplicados en {destination_sheet_name}: {e_del}')

                # Re-leer destino para obtener posiciones actualizadas post-borrado
                dest_resp2 = svc.spreadsheets().values().get(
                    spreadsheetId=consolidado_sheet_id,
                    range=f'{destination_sheet_name}!A:Z'
                ).execute()
                destination_rows = dest_resp2.get('values', [])
                existing_refs_map = {}
                for idx, dest_row in enumerate(destination_rows):
                    if dest_row and len(dest_row) > 0:
                        ref_value = _make_dest_key(dest_row)
                        if ref_value and ref_value not in existing_refs_map:
                            existing_refs_map[ref_value] = idx + 1

        # Deduplicar la fuente usando _make_src_key (posiciones originales, sin tipo).
        # Conservar solo la última ocurrencia de cada clave compuesta.
        seen_source_refs = {}
        for i, row in enumerate(rows_to_copy):
            if row and len(row) > 0:
                ref_value = _make_src_key(row)
                if ref_value:
                    seen_source_refs[ref_value] = i
        deduped_rows_to_copy = []
        for i, row in enumerate(rows_to_copy):
            if row and len(row) > 0:
                ref_value = _make_src_key(row)
                if ref_value:
                    if seen_source_refs.get(ref_value) == i:
                        deduped_rows_to_copy.append(row)
                else:
                    deduped_rows_to_copy.append(row)
            else:
                deduped_rows_to_copy.append(row)
        rows_to_copy = deduped_rows_to_copy

        # Clasificar filas a copiar: nuevas vs a actualizar.
        # La comparación usa _make_src_key (fuente) contra existing_refs_map
        # que fue construido con _make_dest_key — ambas claves son equivalentes
        # porque apuntan a los mismos campos (ref+sku) en sus respectivas filas.
        rows_to_insert = []
        rows_to_update = []  # [(row_number, row_data), ...]

        for row in rows_to_copy:
            if row and len(row) > 0:
                ref_value = _make_src_key(row)
                if ref_value:
                    if ref_value in existing_refs_map:
                        # Ya existe en destino: actualizar
                        row_number = existing_refs_map[ref_value]
                        rows_to_update.append((row_number, row))
                    else:
                        # Nueva: insertar y registrar para no volver a insertar
                        rows_to_insert.append(row)
                        existing_refs_map[ref_value] = -1  # marca como "ya encolada"
                else:
                    # Fila sin ref (posible encabezado sin ID)
                    rows_to_insert.append(row)
            else:
                rows_to_insert.append(row)

        # Prepend tipo en col A justo antes de escribir, ya terminada la clasificación.
        if tipo:
            rows_to_insert = [[tipo] + list(r) for r in rows_to_insert]
            rows_to_update = [(num, [tipo] + list(r)) for num, r in rows_to_update]

        # Procesar actualizaciones de filas duplicadas
        # Agrupar en un solo batch para evitar exceder quotas de write requests
        if rows_to_update:
            batch_data = []
            for row_number, row_data in rows_to_update:
                max_cols = len(row_data)
                end_col = _taxes_col_to_a1(max_cols)
                rng = f'{destination_sheet_name}!A{row_number}:{end_col}{row_number}'
                batch_data.append({'range': rng, 'values': [row_data]})

            # Enviar en chunks para no crear payloads enormes
            def _send_batches(data_list, chunk_size=50, max_retries=6):
                import time, random
                for i in range(0, len(data_list), chunk_size):
                    chunk = data_list[i:i+chunk_size]
                    body = {'valueInputOption': 'USER_ENTERED', 'data': chunk}
                    for attempt in range(max_retries):
                        try:
                            svc.spreadsheets().values().batchUpdate(
                                spreadsheetId=consolidado_sheet_id,
                                body=body
                            ).execute()
                            break
                        except Exception as e:
                            # Detectar 429 en mensaje y aplicar backoff
                            err_str = str(e)
                            if '429' in err_str or 'RATE_LIMIT' in err_str or 'rateLimitExceeded' in err_str or 'RATE_LIMIT_EXCEEDED' in err_str:
                                sleep = min((2 ** attempt) + random.random(), 60)
                                print(f"[taxes_thread] 429 detected, retrying chunk in {sleep:.1f}s (attempt {attempt+1})")
                                time.sleep(sleep)
                                continue
                            else:
                                print(f'[taxes_thread] Error batch-updating filas en {destination_sheet_name}: {e}')
                                break

            try:
                _send_batches(batch_data)
            except Exception as e:
                print(f'[taxes_thread] Error enviando batches de actualización: {e}')

        # Si no hay filas nuevas, retornar con solo actualización
        if not rows_to_insert:
            msg = (
                f"✅ Consolidado actualizado en {destination_sheet_name}: "
                f"{len(rows_to_update)} filas actualizadas desde {source_sheet_name}."
            )
            print(f'[taxes_thread] {msg}')
            if log_widget:
                log_widget.insert('end', msg + '\n')
            return len(rows_to_update)

        rows_to_copy = rows_to_insert

        # Recortar al ancho real para no escribir columnas vacías y consumir
        # menos celdas en el libro destino (Ola 3).
        ancho = 1
        for r in rows_to_copy:
            if r:
                for i in range(len(r) - 1, -1, -1):
                    if str(r[i]).strip():
                        ancho = max(ancho, i + 1)
                        break
        rows_to_copy = [list(r[:ancho]) + [''] * max(0, ancho - len(r)) for r in rows_to_copy]
        end_col = _taxes_col_to_a1(ancho)

        def _append_filas_nuevas(target_sheet_id):
            """Inserta `rows_to_copy` al final de `destination_sheet_name` en
            target_sheet_id usando values().append(), que extiende la grilla
            en forma diferida sin requerir batchUpdate previo. Devuelve
            (filas_escritas, error_msg).
            """
            try:
                # Asegurar que la pestaña destino exista en el target
                meta = svc.spreadsheets().get(spreadsheetId=target_sheet_id).execute()
                tabs_existentes = {sh.get('properties', {}).get('title') for sh in meta.get('sheets', [])}
                if destination_sheet_name not in tabs_existentes:
                    svc.spreadsheets().batchUpdate(
                        spreadsheetId=target_sheet_id,
                        body={'requests': [{'addSheet': {'properties': {'title': destination_sheet_name}}}]}
                    ).execute()
                    print(f'[taxes_thread] Pestaña {destination_sheet_name} creada en {target_sheet_id}.')
            except Exception as e_meta:
                return 0, str(e_meta)

            try:
                svc.spreadsheets().values().append(
                    spreadsheetId=target_sheet_id,
                    range=f'{destination_sheet_name}!A:{end_col}',
                    valueInputOption='USER_ENTERED',
                    insertDataOption='INSERT_ROWS',
                    body={'values': rows_to_copy}
                ).execute()
                return len(rows_to_copy), ''
            except Exception as e_app:
                return 0, str(e_app)

        # Intento primario: appendear en el libro principal
        written, err_primary = _append_filas_nuevas(consolidado_sheet_id)
        target_used = consolidado_sheet_id

        # Si falló por límite de celdas y hay overflow configurado, reintentamos
        es_cell_limit = '10000000 cells' in err_primary or 'limit of 10000000' in err_primary
        if written == 0 and es_cell_limit and overflow_sheet_id:
            msg_fb = (
                f"⚠️ Consolidado primario {consolidado_sheet_id} lleno (10M celdas). "
                f"Reintentando filas nuevas en overflow {overflow_sheet_id}."
            )
            print(f'[taxes_thread] {msg_fb}')
            if log_widget:
                log_widget.insert('end', msg_fb + '\n')
            written, err_overflow = _append_filas_nuevas(overflow_sheet_id)
            if written:
                target_used = overflow_sheet_id

        if written:
            if rows_to_update:
                msg = (
                    f"✅ Consolidado actualizado en {destination_sheet_name}: "
                    f"{len(rows_to_update)} filas actualizadas (en {consolidado_sheet_id}) + "
                    f"{written} nuevas (en {target_used}) desde {source_sheet_name}."
                )
            else:
                msg = (
                    f"✅ Consolidado actualizado en {destination_sheet_name}: "
                    f"{written} filas nuevas en {target_used} desde {source_sheet_name}."
                )
            print(f'[taxes_thread] {msg}')
            if log_widget:
                log_widget.insert('end', msg + '\n')
            return written + len(rows_to_update)

        # FalCourier2 de inserción de filas nuevas
        err_txt = err_primary or 'desconocido'
        if es_cell_limit:
            if overflow_sheet_id:
                hint = (
                    f"❌ No se pudieron insertar {len(rows_to_copy)} filas nuevas en {destination_sheet_name}: "
                    f"primary ({consolidado_sheet_id}) y overflow ({overflow_sheet_id}) fallaron. "
                    f"Detalle primary: {err_txt[:200]}"
                )
            else:
                hint = (
                    f"❌ No se pudieron insertar {len(rows_to_copy)} filas nuevas en {destination_sheet_name}: "
                    f"el libro {consolidado_sheet_id} alcanzó el límite de 10M celdas. "
                    f"Solución: crear un Google Sheet nuevo y setear "
                    f"TAXES_CONSOLIDADO_OVERFLOW_SHEET_ID con su ID."
                )
        else:
            hint = f'⚠️ No se pudieron insertar filas nuevas en {destination_sheet_name}: {err_txt[:300]}'
        print(f'[taxes_thread] {hint}')
        if log_widget:
            log_widget.insert('end', hint + '\n')
        # Devolvemos las que sí actualizamos (dups), aunque las nuevas hayan fallado
        return len(rows_to_update)
    except Exception as e:
        print(f'[taxes_thread] No se pudo consolidar en {destination_sheet_name}: {e}')
        if log_widget:
            log_widget.insert('end', f'⚠️ No se pudo consolidar en {destination_sheet_name}: {e}\n')
        return 0


def taxes_thread():
    global select_option_var,usuario_redshift, contrasena_redshift,usuario_redshift_var,contrasena_redshift_var, seller,fecha_inicio_var, fecha_fin_var, courier_var, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find

    import datetime
    import psycopg2
    import os
    from dotenv import load_dotenv
    load_dotenv()

    # Variables que deben venir del front
    print(usuario_redshift, contrasena_redshift)

    try:
        conn1 = psycopg2.connect(
            host=os.getenv('REDSHIFT_HOST'),
            port=os.getenv('REDSHIFT_PORT'),
            dbname=os.getenv('REDSHIFT_DB'),
            user=usuario_redshift,
            password=contrasena_redshift,
            connect_timeout=5
        )
        conn1.close()
        print("[taxes_thread] Redshift credentials OK")
    except Exception as auth_err:
        msg = f"⚠️ Error conectando a Redshift: {auth_err}"
        print(msg)
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    fecha_inicio = fecha_inicio_var.get()
    fecha_fin = fecha_fin_var.get()
    courier = courier_var.get() if 'courier_var' in globals() else None
    print(f"[taxes_thread] Fecha inicio seleccionada: {fecha_inicio}")
    print(f"[taxes_thread] Fecha fin seleccionada: {fecha_fin}")
    print(f"[taxes_thread] Courier seleccionado: {courier}")

    # Validaciones de fechas y courier
    if not fecha_inicio or not fecha_fin:
        msg = "⚠️ Debe seleccionar ambas fechas."
        print(f"[taxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return
    try:
        fecha_inicio_dt = datetime.datetime.strptime(fecha_inicio, "%Y-%m-%d")
        fecha_fin_dt = datetime.datetime.strptime(fecha_fin, "%Y-%m-%d")
    except Exception as e:
        msg = f"⚠️ Formato de fecha incorrecto: {e}"
        print(f"[taxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return
    if fecha_inicio_dt > fecha_fin_dt:
        msg = "⚠️ La fecha de inicio no puede ser mayor que la fecha de fin."
        print(f"[taxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return
    if not courier or courier.strip() == "":
        msg = "⚠️ Debe seleccionar un courier."
        print(f"[taxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    # ID de la carpeta de Google Drive y del Google Sheet destino
    if courier == "Courier1":
        DRIVE_FOLDER_ID = "1Ztsv1c9oqE8rfHdyDHjECDa2qKplBtre"
        SHEET_ID = "1D0YFVnWzbsnmi5gZNpJ6XVENgC5Blo116OpQ-evggzc"
        SHEET_NAME = "INPUT Courier"
        SHEET_TM = "INPUT TM - Pos"
        SHEET_SOURCE = SHEET_TM
        pais = "Argentina"
    elif courier == "Courier2":
        DRIVE_FOLDER_ID = "1yV8VTpz_Q6Z7-FSmaAFcHl-HPfze5iO7"
        SHEET_ID = "1oy0bXnZGzknb91jFwXmtc_QGx0tyzY7AcVc19iRUHNc"
        SHEET_NAME = "INPUT Courier"
        SHEET_TM = "INPUT TM - Pos"
        SHEET_SOURCE = SHEET_TM
        pais = "Costa Rica"
    elif courier == "Courier3":
        SHAREPOINT_URL = "https://Courier3courierpe-my.sharepoint.com/:x:/g/personal/aarias_Courier3courierpe_onmicrosoft_com/IQCo00dS99tzRomsWwTtljLxrAWdQFxCGq6vW_c_LGUDIw5s?download=1"
        SHEET_ID = "1QNsX-_j-_hdiGkf79czaLCkePX4x5n3rmN0bxK7uu7I"
        SHEET_NAME = "INPUT Courier"
        SHEET_SOURCE = SHEET_NAME
        pais = "Peru"
    elif courier == "Courier4":
        DRIVE_FOLDER_ID = "1UIY85aNoMYOj2KNPt-53FGcA9L3Xuyxt"
        SHEET_ID = "1FW71FZ2lBktZuUOSjWQnOChLpA7JCzBhmDv9uqLRhMM"
        SHEET_NAME = "INPUT Courier"
        SHEET_TM = "INPUT TM - Pos"
        SHEET_SOURCE = SHEET_TM
        pais = "Ecuador"
    elif courier == "Courier5":
        DRIVE_FOLDER_ID = "1-T9E9Bg-ZpNJwSgye5efnW6jNaLWFODo"
        SHEET_ID = "1b8tTL-0Z7DZEER-ewtNp4O7NpWZLqchTHFmLu7sT0eE"
        SHEET_NAME = "INPUT Courier"
        SHEET_TM = "INPUT TM - Pos"
        SHEET_SOURCE = SHEET_TM
        pais = "Uruguay"
    else:
        msg = f"⚠️ Courier no soportado: {courier}"
        print(f"[taxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return


    CRED_FILE = "credenciales_drive.json"

    import io
    import pandas as pd
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2.service_account import Credentials
    import datetime


    ordenes = []
    try:
        print("[taxes_thread] Iniciando autenticación y búsqueda de archivos...")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', "🔑 Autenticando con Google...\n")
        scopes = ["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/spreadsheets"]
        credentials = Credentials.from_service_account_file(CRED_FILE, scopes=scopes)
        drive_service = build("drive", "v3", credentials=credentials)
        sheets_service = build("sheets", "v4", credentials=credentials)

        def _fallback_Courier5_sheets_service(error_obj, contexto):
            """Para Courier5, si falla por permisos en Sheets con credenciales_drive, reintenta con credenciales_drive."""
            nonlocal sheets_service
            err_txt = str(error_obj or "")
            if courier != "Courier5":
                raise error_obj
            if "403" not in err_txt and "does not have permission" not in err_txt.lower():
                raise error_obj

            try:
                alt_scopes = [
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
                alt_credentials = Credentials.from_service_account_file("credenciales_drive.json", scopes=alt_scopes)
                sheets_service = build("sheets", "v4", credentials=alt_credentials)
                print(f"[taxes_thread] Reintentando {contexto} con credenciales_drive.json por permisos 403...")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', f"🔁 Reintento de {contexto} con credenciales_drive.json por permisos...\\n")
            except Exception as fallback_err:
                print(f"[taxes_thread] No se pudo activar fallback de Sheets: {fallback_err}")
                raise error_obj

        # Para Courier3: descarga desde SharePoint y sube directo a Sheets (sin Redshift ni IA)
        if courier == "Courier3":
            _taxes_Courier3_procesar(
                SHAREPOINT_URL, SHEET_ID, SHEET_NAME,
                fecha_inicio_dt, fecha_fin_dt,
                sheets_service, log_text, ordenes
            )
            _taxes_consolidar_input_tm_a_historial(
                sheets_service,
                SHEET_ID,
                SHEET_SOURCE,
                courier,
                log_text,
                key_cols=(0,),
                tipo="Pos",
            )
            return

        # Para Courier4: procesar PDFs
        if courier == "Courier4":
            datos = _process_Courier4_pdfs(drive_service, DRIVE_FOLDER_ID, fecha_inicio_dt, fecha_fin_dt, log_text)
            if not datos:
                log_text.insert('end', "No se encontraron registros válidos\n")
                return

            import pandas as pd
            df_final = pd.DataFrame({
                0: [d['identificador'] for d in datos],
                1: [""] * len(datos),
                2: [""] * len(datos),
                3: [""] * len(datos),
                4: [d['total_usd'] for d in datos],
                5: [""] * len(datos),
            })

            # Subir a Sheets
            clear_range = f"{SHEET_NAME}!A2:F"
            sheets_service.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range=clear_range).execute()

            import math
            def convert_value(obj):
                if isinstance(obj, float) and math.isnan(obj):
                    return ""
                if isinstance(obj, __import__('decimal').Decimal):
                    return float(obj)
                return obj

            values = [[convert_value(c) for c in r] for r in df_final.values.tolist()]
            sr = 2
            er = sr + len(values) - 1
            ce = chr(ord('A') + df_final.shape[1] - 1)
            rd = f"{SHEET_NAME}!A{sr}:{ce}{er}"

            sheets_service.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=rd, valueInputOption="USER_ENTERED", body={"values": values}).execute()

            log_text.insert('end', f"✅ {len(values)} registros subidos a INPUT Courier\n")

            # Búsqueda intermedia: combinar órdenes directas con las heredadas por manifiesto
            log_text.insert('end', "🔍 Leyendo identificadores desde INPUT Courier...\n")
            manifests_resp = sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=f"{SHEET_NAME}!A2:A"
            ).execute()
            manifests_rows = manifests_resp.get('values', [])

            ordenes_directas = []
            vuelos_manifests = []
            seen_directas = set()
            seen_manifests = set()
            for row in manifests_rows:
                raw_value = str(row[0]).strip() if row and len(row) > 0 else ""
                normalized = ' '.join(raw_value.upper().split())
                if not normalized:
                    continue
                if normalized.startswith("Courier4"):
                    if normalized in seen_manifests:
                        continue
                    seen_manifests.add(normalized)
                    vuelos_manifests.append(normalized)
                else:
                    if normalized in seen_directas:
                        continue
                    seen_directas.add(normalized)
                    ordenes_directas.append(normalized)

            log_text.insert('end', f"📄 Identificadores directos leídos: {len(ordenes_directas)} | manifiestos heredados: {len(vuelos_manifests)}\n")

            ordenes = list(ordenes_directas)

            if vuelos_manifests:
                try:
                    conn_temp = psycopg2.connect(
                        host=os.getenv('REDSHIFT_HOST'),
                        port=os.getenv('REDSHIFT_PORT'),
                        dbname=os.getenv('REDSHIFT_DB'),
                        user=usuario_redshift,
                        password=contrasena_redshift
                    )
                    cur_temp = conn_temp.cursor()

                    patterns = [f"{manifest}%" for manifest in vuelos_manifests]
                    where_like = " OR ".join(["UPPER(custom_manifest_number) LIKE %s"] * len(patterns))
                    query_manifest = (
                        "SELECT DISTINCT ref_id "
                        "FROM tm_staging.brightpearl_sales_orders "
                        "WHERE ref_id IS NOT NULL AND (" + where_like + ")"
                    )
                    print(f"[taxes_thread] Consultando manifests Courier4 únicos: {len(vuelos_manifests)}")
                    cur_temp.execute(query_manifest, patterns)
                    rows = cur_temp.fetchall()
                    ordenes_desde_manifiestos = sorted({str(row[0]).strip() for row in rows if row and row[0]})

                    cur_temp.close()
                    conn_temp.close()

                    if ordenes_desde_manifiestos:
                        ordenes = sorted(set(ordenes + ordenes_desde_manifiestos))

                    if not ordenes:
                        log_text.insert('end', "⚠️ No se encontraron órdenes para estos archivos\n")
                        return

                    log_text.insert('end', f"✅ {len(ordenes)} órdenes encontradas\n")

                except Exception as e:
                    log_text.insert('end', f"❌ Error buscando órdenes: {str(e)[:100]}\n")
                    print(f"[taxes_thread] Error en búsqueda de manifests: {e}")
                    return

            if not ordenes:
                log_text.insert('end', "⚠️ No se encontraron órdenes para consultar en Redshift\n")
                return

            # Continuar a Redshift con las órdenes reales

        # Para Courier5 (Uruguay): leer el número de carga y buscar ref_id por custom_manifest_number
        if courier == "Courier5":
            query = (
                f"'{DRIVE_FOLDER_ID}' in parents and ("
                "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or "
                "mimeType='application/vnd.ms-excel' or "
                "mimeType='application/vnd.ms-excel.sheet.macroEnabled.12' or "
                "mimeType='application/msword' or "
                "mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"
                ") and trashed = false"
            )
            print(f"[taxes_thread] Query de búsqueda Courier5: {query}")
            results = drive_service.files().list(
                q=query,
                fields="files(id, name, modifiedTime)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files = results.get('files', [])
            print(f"[taxes_thread] Archivos encontrados en Drive (Courier5): {len(files)}")

            archivos_filtrados = []
            for f in files:
                mod_time = datetime.datetime.strptime(f['modifiedTime'][:10], "%Y-%m-%d")
                if fecha_inicio_dt <= mod_time <= fecha_fin_dt:
                    archivos_filtrados.append(f)

            print(f"[taxes_thread] Archivos Courier5 filtrados por fecha: {len(archivos_filtrados)}")
            if not archivos_filtrados:
                msg = "No se encontraron archivos Courier5 en el rango de fechas."
                print(f"[taxes_thread] {msg}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
                return

            def _normalizar_monto_uy(valor_raw):
                txt = str(valor_raw or '').strip()
                if not txt:
                    return None
                # Formato esperado: 234.664,00 -> 234664.00
                txt = txt.replace(' ', '').replace('$', '')
                txt = txt.replace('.', '').replace(',', '.')
                try:
                    return float(txt)
                except Exception:
                    return None

            def _extraer_cargas_y_totales_Courier5(texto):
                carga_a_total = {}
                if not texto:
                    return carga_a_total

                upper_text = str(texto).upper()

                # Cargas posibles en el texto
                cargas = []
                for patron in [
                    r"Courier5\s+IMPORTACIONES\s*(\d{2,6})",
                    r"IMPORTACIONES\s*(\d{2,6})",
                    r"\bTM\s*(\d{2,6})\b",
                ]:
                    for m in re.finditer(patron, upper_text):
                        try:
                            cargas.append((str(int(m.group(1))), m.start()))
                        except Exception:
                            pass

                # Totales posibles: "Total.................$    234.664,00"
                total_matches = []
                patron_total = r"TOTAL[^\d$]{0,40}\$\s*([0-9\.,]{3,})"
                for m in re.finditer(patron_total, upper_text):
                    monto = _normalizar_monto_uy(m.group(1))
                    if monto is not None:
                        total_matches.append((monto, m.start()))

                if not cargas:
                    return carga_a_total

                # Si hay una sola carga y un solo total, asignación directa
                if len(cargas) == 1 and len(total_matches) == 1:
                    carga_a_total[cargas[0][0]] = total_matches[0][0]
                    return carga_a_total

                # Si hay varias cargas pero un único total, usar ese total para todas
                if len(total_matches) == 1:
                    monto_unico = total_matches[0][0]
                    for carga, _ in cargas:
                        carga_a_total[carga] = monto_unico
                    return carga_a_total

                # Emparejamiento por cercanía posicional (fallback)
                for carga, pos_carga in cargas:
                    if not total_matches:
                        carga_a_total.setdefault(carga, None)
                        continue
                    monto_cercano = min(total_matches, key=lambda t: abs(t[1] - pos_carga))[0]
                    carga_a_total[carga] = monto_cercano

                return carga_a_total

            cargas_detectadas = set()
            carga_total_map = {}
            for archivo in archivos_filtrados:
                nombre = str(archivo.get('name', '')).strip()
                map_nombre = _extraer_cargas_y_totales_Courier5(nombre)
                for carga, total in map_nombre.items():
                    cargas_detectadas.add(carga)
                    if total is not None and carga not in carga_total_map:
                        carga_total_map[carga] = total

                # Intentar leer contenido del archivo para detectar carga + total
                try:
                    request = drive_service.files().get_media(fileId=archivo['id'])
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                    raw_bytes = fh.getvalue()
                    texto = raw_bytes.decode('utf-8', errors='ignore')
                    if not texto.strip():
                        texto = raw_bytes.decode('latin-1', errors='ignore')

                    map_contenido = _extraer_cargas_y_totales_Courier5(texto)
                    for carga, total in map_contenido.items():
                        cargas_detectadas.add(carga)
                        if total is not None:
                            carga_total_map[carga] = total
                except Exception as e:
                    print(f"[taxes_thread] Aviso: no se pudo leer contenido de {nombre}: {e}")

            if not cargas_detectadas:
                msg = "⚠️ No se pudo detectar el número de carga Courier5 en los archivos seleccionados."
                print(f"[taxes_thread] {msg}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
                return

            manifiestos_Courier5 = [f"Courier5 IMPORTACIONES {carga}" for carga in sorted(cargas_detectadas, key=lambda x: int(x))]
            print(f"[taxes_thread] Manifiestos Courier5 detectados: {manifiestos_Courier5}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', f"📄 Cargas Courier5 detectadas: {', '.join(manifiestos_Courier5)}\n")

            # Guardar manifiestos detectados en INPUT Courier (columna A)
            clear_range = f"{SHEET_NAME}!A2:F"
            try:
                sheets_service.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range=clear_range).execute()
            except Exception as e:
                _fallback_Courier5_sheets_service(e, "limpieza de INPUT Courier")
                sheets_service.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range=clear_range).execute()
            values_Courier5 = []
            for m in manifiestos_Courier5:
                carga = m.replace("Courier5 IMPORTACIONES", "").strip()
                total_carga = carga_total_map.get(carga)
                values_Courier5.append([m, "", "", "", total_carga if total_carga is not None else "", ""])

            totales_log = [
                f"{m}=${carga_total_map.get(m.replace('Courier5 IMPORTACIONES', '').strip())}"
                for m in manifiestos_Courier5
                if carga_total_map.get(m.replace('Courier5 IMPORTACIONES', '').strip()) is not None
            ]
            if totales_log:
                print(f"[taxes_thread] Totales Courier5 detectados: {totales_log}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', f"💰 Totales detectados: {', '.join(totales_log)}\n")
            sr = 2
            er = sr + len(values_Courier5) - 1
            rd = f"{SHEET_NAME}!A{sr}:F{er}"
            try:
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SHEET_ID,
                    range=rd,
                    valueInputOption="USER_ENTERED",
                    body={"values": values_Courier5}
                ).execute()
            except Exception as e:
                _fallback_Courier5_sheets_service(e, "escritura en INPUT Courier")
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SHEET_ID,
                    range=rd,
                    valueInputOption="USER_ENTERED",
                    body={"values": values_Courier5}
                ).execute()

            # Buscar órdenes por custom_manifest_number
            log_text.insert('end', "🔍 Buscando órdenes Courier5 por número de carga...\n")
            try:
                conn_temp = psycopg2.connect(
                    host=os.getenv('REDSHIFT_HOST'),
                    port=os.getenv('REDSHIFT_PORT'),
                    dbname=os.getenv('REDSHIFT_DB'),
                    user=usuario_redshift,
                    password=contrasena_redshift
                )
                cur_temp = conn_temp.cursor()

                patterns = [f"{manifest}%" for manifest in manifiestos_Courier5]
                where_like = " OR ".join(["UPPER(custom_manifest_number) LIKE %s"] * len(patterns))
                query_manifest = (
                    "SELECT DISTINCT ref_id "
                    "FROM tm_staging.brightpearl_sales_orders "
                    "WHERE ref_id IS NOT NULL AND (" + where_like + ")"
                )
                print(f"[taxes_thread] Consultando manifests Courier5 únicos: {len(manifiestos_Courier5)}")
                cur_temp.execute(query_manifest, patterns)
                rows_Courier5 = cur_temp.fetchall()
                ordenes = sorted({str(row[0]).strip() for row in rows_Courier5 if row and row[0]})

                cur_temp.close()
                conn_temp.close()

                if not ordenes:
                    log_text.insert('end', "⚠️ No se encontraron órdenes para las cargas Courier5 detectadas\n")
                    return

                log_text.insert('end', f"✅ {len(ordenes)} órdenes Courier5 encontradas\n")

            except Exception as e:
                log_text.insert('end', f"❌ Error buscando órdenes Courier5: {str(e)[:100]}\n")
                print(f"[taxes_thread] Error en búsqueda de manifests Courier5: {e}")
                return

        if courier not in ("Courier4", "Courier5"):
            # Buscar archivos Excel en la carpeta de Drive
            query = (
                f"'{DRIVE_FOLDER_ID}' in parents and ("
                "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or "
                "mimeType='application/vnd.ms-excel' or "
                "mimeType='application/vnd.ms-excel.sheet.macroEnabled.12'"
                ") and trashed = false"
            )
            print(f"[taxes_thread] Query de búsqueda: {query}")
            results = drive_service.files().list(q=query, fields="files(id, name, modifiedTime)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            files = results.get('files', [])
            print(f"[taxes_thread] Archivos encontrados en Drive: {len(files)}")

            # Filtrar archivos por fecha
            archivos_filtrados = []
            for f in files:
                mod_time = datetime.datetime.strptime(f['modifiedTime'][:10], "%Y-%m-%d")
                if fecha_inicio_dt <= mod_time <= fecha_fin_dt:
                    archivos_filtrados.append(f)

            print(f"[taxes_thread] Archivos filtrados por fecha: {len(archivos_filtrados)}")
            if not archivos_filtrados:
                msg = "No se encontraron archivos en el rango de fechas."
                print(f"[taxes_thread] {msg}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
                return

            # Descargar y unir los datos de todos los archivos Excel

            dfs = []
            log_text.insert('end', "⬇️ Descargando y leyendo archivos Excel...\n")
            for archivo in archivos_filtrados:
                print(f"[taxes_thread] Descargando y leyendo archivo: {archivo['name']}")
                request = drive_service.files().get_media(fileId=archivo['id'])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                try:
                    def _parse_total_terminal(value):
                        if pd.isna(value):
                            return None
                        txt = str(value).strip()
                        if not txt:
                            return None
                        txt = txt.replace("$", "").replace(" ", "").replace(",", "")
                        if txt in {"-", "", "nan", "None"}:
                            return None
                        try:
                            return float(txt)
                        except Exception:
                            return None

                    def _parse_cif_terminal(value):
                        if pd.isna(value):
                            return None
                        txt = str(value).strip()
                        if not txt:
                            return None
                        txt = txt.replace("$", "").replace(" ", "").replace(",", "")
                        if txt in {"-", "", "nan", "None"}:
                            return None
                        try:
                            return float(txt)
                        except Exception:
                            return None

                    def _normalize_order_terminal(value):
                        # Filtro defensivo: nulos / vacíos
                        try:
                            if pd.isna(value):
                                return ""
                        except (TypeError, ValueError):
                            pass

                        # Si viene como número (float/int), pasar por int para evitar
                        # que un valor leído como float (p.ej. 268806.0) deje un "0"
                        # extra al quedarnos sólo con los dígitos.
                        import math
                        if isinstance(value, bool):
                            return ""
                        if isinstance(value, (int, float)):
                            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                                return ""
                            try:
                                value = int(value)
                            except (ValueError, OverflowError):
                                return ""

                        txt = str(value).strip()
                        if not txt:
                            return ""

                        # Si el string es un número con parte decimal (ej: "268806.0"
                        # o "268806,00"), quedarnos sólo con la parte entera para no
                        # arrastrar dígitos sobrantes del decimal.
                        import re
                        m = re.match(r'^\s*([+-]?\d+)[.,]\d+\s*$', txt)
                        if m:
                            txt = m.group(1)

                        txt_upper = txt.upper()
                        if txt_upper.startswith("CR"):
                            suffix = ''.join(ch for ch in txt_upper[2:] if ch.isdigit())
                            if not suffix:
                                return ""
                            # Validación: el número de orden CR no debería superar 9 dígitos.
                            # Si excede, dejamos los últimos 9 (lo más probable es que el sobrante
                            # provenga de un decimal mal formateado o un sufijo extra).
                            if len(suffix) > 9:
                                print(f"[taxes_thread] Aviso: orden con más de 9 dígitos ({suffix!r}); se recortan los últimos 9")
                                suffix = suffix[-9:]
                            return f"CR{suffix.zfill(9)}"
                        digits = ''.join(ch for ch in txt if ch.isdigit())
                        if not digits:
                            return ""
                        if len(digits) > 9:
                            print(f"[taxes_thread] Aviso: orden con más de 9 dígitos ({digits!r}); se recortan los últimos 9")
                            digits = digits[-9:]
                        return f"CR{digits.zfill(9)}"

                    def _leer_terminal_individual(buffer):
                        buffer.seek(0)
                        sheet_names = pd.ExcelFile(buffer).sheet_names
                        matched = next(
                            (s for s in sheet_names if "INDIVIDUAL" in s.upper()), None
                        )
                        if matched is None:
                            raise ValueError(
                                f"No se encontró hoja con 'INDIVIDUAL' en el archivo. Hojas disponibles: {sheet_names}"
                            )
                        buffer.seek(0)
                        df_raw = pd.read_excel(
                            buffer,
                            sheet_name=matched,
                            header=None
                        )

                        header_idx = None
                        guia_idx = None
                        cif_idx = None
                        total_idx = None
                        for i, row_vals in df_raw.iterrows():
                            valores = [str(v).strip().upper() if pd.notna(v) else "" for v in row_vals.tolist()]
                            if "FECHA" in valores and "GUIA" in valores and "CIF" in valores and "TOTAL" in valores:
                                header_idx = i
                                guia_idx = valores.index("GUIA")
                                cif_idx = valores.index("CIF")
                                total_idx = valores.index("TOTAL")
                                break

                        if header_idx is None:
                            raise ValueError(
                                "No se encontró encabezado FECHA/GUIA/CIF/TOTAL en hoja FACTURADOR IMPUESTOS INDIVIDUAL"
                            )

                        df_data = df_raw.iloc[header_idx + 1:, [guia_idx, cif_idx, total_idx]].copy()
                        df_data.columns = ["GUIA", "CIF", "TOTAL"]
                        df_data.loc[:, "GUIA"] = df_data["GUIA"].apply(_normalize_order_terminal)
                        df_data.loc[:, "CIF"] = df_data["CIF"].apply(_parse_cif_terminal)
                        df_data.loc[:, "TOTAL"] = df_data["TOTAL"].apply(_parse_total_terminal)
                        df_data = df_data[(df_data["GUIA"] != "") & (df_data["TOTAL"].notna())].copy()
                        return df_data[["GUIA", "CIF", "TOTAL"]]

                    def _leer_terminal_global(buffer):
                        import warnings
                        buffer.seek(0)
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore")
                            sheet_names = pd.ExcelFile(buffer, engine="openpyxl").sheet_names
                        matched = next(
                            (s for s in sheet_names if "GLOBALES" in s.upper()), None
                        )
                        if matched is None:
                            raise ValueError(
                                f"No se encontró hoja con 'GLOBALES' en el archivo. Hojas disponibles: {sheet_names}"
                            )
                        buffer.seek(0)
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore")
                            df_raw = pd.read_excel(
                                buffer,
                                sheet_name=matched,
                                header=9,
                                engine="openpyxl"
                            )

                        columnas = {str(col).strip().upper(): col for col in df_raw.columns}
                        guia_col = columnas.get("GUIA")
                        cif_col = columnas.get("CIF")
                        total_col = columnas.get("TOTAL")

                        if guia_col is None or cif_col is None or total_col is None:
                            raise ValueError(
                                "No se encontraron las columnas GUIA, CIF y TOTAL en hoja FACTURADOR IMPUESTOS GLOBALES"
                            )

                        df_data = df_raw[[guia_col, cif_col, total_col]].copy()
                        df_data.columns = ["GUIA", "CIF", "TOTAL"]
                        df_data["GUIA"] = df_data["GUIA"].astype(object).apply(_normalize_order_terminal)
                        df_data.loc[:, "CIF"] = df_data["CIF"].apply(_parse_cif_terminal)
                        df_data.loc[:, "TOTAL"] = df_data["TOTAL"].apply(_parse_total_terminal)
                        df_data = df_data[(df_data["GUIA"] != "") & (df_data["TOTAL"].notna())].copy()
                        return df_data

                    if courier == "Courier2":
                        df_parts = [_leer_terminal_individual(fh)]
                        try:
                            if hasattr(fh, "seek"):
                                fh.seek(0)
                            df_global = [_leer_terminal_global(fh)]
                            print(f"[DEBUG] df_global shape: {df_global[0].shape}, empty: {df_global[0].empty}")
                            if not df_global[0].empty:
                                df_parts.append(df_global[0])
                            else:
                                print("[DEBUG] df_global está vacío, no se agrega")
                        except Exception as global_err:
                            import traceback
                            print(f"[taxes_thread] Error en GLOBAL: {global_err}")
                            traceback.print_exc()

                        df_data = pd.concat(df_parts, ignore_index=True)
                        df_data["CIF"] = pd.to_numeric(df_data["CIF"], errors="coerce")
                        df_data["TOTAL"] = pd.to_numeric(df_data["TOTAL"], errors="coerce")
                        df_data = df_data.dropna(subset=["GUIA", "TOTAL"]).copy()
                        df_data = (
                            df_data.groupby("GUIA", as_index=False, sort=False)[["CIF", "TOTAL"]]
                            .sum()
                        )

                        # Respetar formato INPUT Courier:
                        # A Nro.Guia, B Mercaderia(vacio), C Peso(vacio), D CIF, E TotalUSD, F CountryCurrency(vacio)
                        df = pd.DataFrame({
                            0: df_data["GUIA"],
                            1: "",
                            2: "",
                            3: df_data["CIF"],
                            4: df_data["TOTAL"],
                            5: "",
                        })
                    elif courier == "Courier1":
                        # Leer formato de Courier1: A(Nro.Guia), B(Mercaderia), C(Peso), D(FOB), J(Total USD), K(Total ARS)
                        df_raw = pd.read_excel(
                            fh,
                            header=None,
                            usecols="A,B,C,D,J,K"
                        )
                        
                        # Asignar nombres a columnas
                        df_raw.columns = ["Nro_Guia", "Mercaderia", "Peso", "FOB", "Total_USD", "Total_ARS"]
                        
                        # Saltar filas vacías al inicio y encontrar encabezados
                        header_idx = None
                        for i, row_vals in df_raw.iterrows():
                            c0 = str(row_vals.iloc[0]).strip().upper() if pd.notna(row_vals.iloc[0]) else ""
                            if c0 and c0 != "NAN":
                                header_idx = i
                                break
                        
                        if header_idx is None:
                            header_idx = 0
                        
                        # Obtener datos desde la fila encontrada
                        df_data = df_raw.iloc[header_idx + 1:].copy()
                        df_data = df_data.dropna(subset=["Nro_Guia"], how="all")
                        
                        # Limpiar y convertir valores
                        def _parse_numero(value):
                            if pd.isna(value):
                                return ""
                            return str(value).strip()
                        
                        def _parse_numero_float(value):
                            if pd.isna(value):
                                return None
                            txt = str(value).strip()
                            if not txt:
                                return None
                            txt = txt.replace("$", "").replace(" ", "").replace(",", ".")
                            try:
                                return float(txt)
                            except Exception:
                                return None
                        
                        df_data["Nro_Guia"] = df_data["Nro_Guia"].apply(_parse_numero)
                        df_data["Mercaderia"] = df_data["Mercaderia"].apply(_parse_numero)
                        df_data["Peso"] = df_data["Peso"].apply(_parse_numero)
                        df_data["FOB"] = df_data["FOB"].apply(_parse_numero_float)
                        df_data["Total_USD"] = df_data["Total_USD"].apply(_parse_numero_float)
                        df_data["Total_ARS"] = df_data["Total_ARS"].apply(_parse_numero_float)
                        
                        # Formato INPUT Courier: A(Nro.Guia), B(Mercaderia), C(Peso), D(FOB), E(Total USD), F(Total ARS)
                        df = pd.DataFrame({
                            0: df_data["Nro_Guia"],
                            1: df_data["Mercaderia"],
                            2: df_data["Peso"],
                            3: df_data["FOB"],
                            4: df_data["Total_USD"],
                            5: df_data["Total_ARS"],
                        })
                    else:
                        df = pd.read_excel(fh)
                        # Filtrar solo columnas A-F (primeras 6 columnas)
                        df = df.iloc[:, :6]
                    dfs.append(df)
                    print(f"[taxes_thread] Archivo leído correctamente: {archivo['name']}")
                except Exception as e:
                    print(f"[taxes_thread] Error leyendo {archivo['name']}: {e}")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"Error leyendo {archivo['name']}: {e}\n")

            if not dfs:
                msg = "No se pudo leer ningún archivo Excel."
                print(f"[taxes_thread] {msg}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
                return

            df_final = pd.concat(dfs, ignore_index=True)
            print(f"[taxes_thread] Total de filas a subir: {len(df_final)}")
            if df_final.empty:
                msg = "No hay filas válidas para subir en el rango seleccionado."
                print(f"[taxes_thread] {msg}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', msg + '\n')
                return

            # Limpiar datos viejos de la columna A a F (excepto encabezado)
            clear_range = f"{SHEET_NAME}!A2:F"
            print(f"[taxes_thread] Limpiando rango: {clear_range}")
            sheets_service.spreadsheets().values().clear(
                spreadsheetId=SHEET_ID,
                range=clear_range
            ).execute()

            # Subir SOLO los datos (sin encabezado) a la columna A, desde la fila 2 en adelante
            import math

            def convert_value(obj):
                if isinstance(obj, float) and math.isnan(obj):
                    return ""
                if isinstance(obj, decimal.Decimal):
                    return float(obj)
                return obj

            values = [[convert_value(cell) for cell in row] for row in df_final.values.tolist()]
            start_row = 2
            end_row = start_row + len(values) - 1
            col_end = chr(ord('A') + df_final.shape[1] - 1)
            range_destino = f"{SHEET_NAME}!A{start_row}:{col_end}{end_row}"
            print(f"[taxes_thread] Subiendo datos al rango: {range_destino}")
            sheets_service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID,
                range=range_destino,
                valueInputOption="USER_ENTERED",  # Permite que Google Sheets interprete números automáticamente
                body={"values": values}
            ).execute()

            # Formatear columna E como número (FLOAT)
            print(f"[taxes_thread] Formateando columna E como FLOAT...")
            format_range = f"{SHEET_NAME}!E{start_row}:E{end_row}"
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheets_service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()['sheets'][0]['properties']['sheetId'] if SHEET_NAME == "INPUT Courier" else None,
                        "startRowIndex": start_row - 1,
                        "endRowIndex": end_row,
                        "startColumnIndex": 4,
                        "endColumnIndex": 5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "0.00"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            }]
        
            # Obtener el sheetId correcto
            sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
            sheet_id_target = None
            for sheet in sheet_metadata['sheets']:
                if sheet['properties']['title'] == SHEET_NAME:
                    sheet_id_target = sheet['properties']['sheetId']
                    break
        
            if sheet_id_target is not None:
                requests[0]['repeatCell']['range']['sheetId'] = sheet_id_target
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=SHEET_ID,
                    body={"requests": requests}
                ).execute()
                print(f"[taxes_thread] Columna E formateada como FLOAT exitosamente.")
            else:
                print(f"[taxes_thread] ⚠️ No se pudo encontrar el sheetId para {SHEET_NAME}")

            msg = f"✅ {len(archivos_filtrados)} archivos procesados y datos subidos a Google Sheet desde la fila 2."
            print(f"[taxes_thread] {msg}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', msg + '\n')

        # --- NUEVO: Extraer columna A (números de orden) ---
        # Para Courier4/Courier5, ya tenemos órdenes de la búsqueda intermedia
        if courier not in ("Courier4", "Courier5"):
            print("[taxes_thread] Extrayendo números de orden de la columna A...")
            ordenes = [str(row[0]).strip() for row in values if str(row[0]).strip()]
            print(f"[taxes_thread] Total de órdenes encontradas: {len(ordenes)}")
            if not ordenes:
                print("[taxes_thread] No se encontraron órdenes para consultar en Redshift.")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', "No se encontraron órdenes para consultar en Redshift.\n")
                return

        # --- NUEVO: Consultar Redshift ---
        print("[taxes_thread] Consultando Redshift...")
        log_text.insert('end', "🔑 Conectando a Redshift...\n")
        try:
            print(os.getenv('REDSHIFT_HOST'), usuario_redshift, contrasena_redshift)
            print(f"Total de órdenes a consultar: {len(ordenes)}")
            print(f"Primeras órdenes: {ordenes[:5]}")
            if not ordenes:
                print("No hay órdenes para consultar en Redshift.")
                return
            
            conn1 = psycopg2.connect(
                host=os.getenv('REDSHIFT_HOST'),
                port=os.getenv('REDSHIFT_PORT'),
                dbname=os.getenv('REDSHIFT_DB'),
                user=usuario_redshift,
                password=contrasena_redshift
            )
            cur1 = conn1.cursor()
            # Armar la lista de órdenes para la query
            ordenes_sql = ','.join([f"'{o}'" for o in ordenes])
            query = f"""
WITH order_items AS (
    SELECT
         omsoi.ref_id         AS ref
        ,omsoi.item_sku       AS sku
        ,UPPER(REPLACE(SUBSTRING(omsoi.item_sku, CHARINDEX('-', omsoi.item_sku) + 1), '_', '')) AS clean_sku
        ,omsoi.item_cost
        ,omsoi.item_net       AS fob_item
        ,omsoi.totalvalue_net AS fob_orden
        ,omsoi.item_name      AS name
        ,omsoi.quantity
        ,omsoi.updated_at_oms
    FROM tm_staging.brightpearl_purchase_order_item AS omsoi
    WHERE omsoi.ref_id IN ({ordenes_sql})
        AND omsoi.item_sku IS NOT NULL
        AND omsoi.item_sku <> ''
        AND omsoi.statusid_oms IN ('1.1 Placed with supplier', '3.1 Invoice received')
), sales_orders AS (
    SELECT DISTINCT
         ref_id
        ,custom_manifest_number
    FROM tm_staging.brightpearl_sales_orders
    WHERE ref_id IN (SELECT ref FROM order_items)
), manifest_data AS (
    SELECT
         mdi.manifest_data_order_mgt AS ref
        ,mdi.sku                     AS sku
        ,UPPER(REPLACE(REPLACE(SUBSTRING(mdi.sku, CHARINDEX('-', mdi.sku) + 1), '_', ''), '|', '')) AS clean_sku
        ,mdi.category                AS categoria
        ,mdi.total_weight
        ,mdi.quantity_real
        ,mdi.price
        ,ROW_NUMBER() OVER (
            PARTITION BY mdi.manifest_data_order_mgt, mdi.sku
            ORDER BY mdi.manifest_data_order_mgt DESC
        ) AS rn
    FROM tm_staging.m1_manifest_data_items_delsert AS mdi
    INNER JOIN order_items AS oi
        ON mdi.manifest_data_order_mgt = oi.ref
    WHERE mdi.url <> 'non-image'
), tax_raw AS (
    SELECT increment_id AS ref, pricing_data AS p_data
    FROM tm_staging.m2_sales_order_delsert
    WHERE increment_id IN (SELECT ref FROM order_items)
), tax_income_item AS (
    SELECT
         b.ref
        ,TRIM('"' FROM CAST(i.sku AS VARCHAR)) AS sku
        ,MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'tax' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS tax_total
    FROM tax_raw AS b
    LEFT JOIN b.p_data.carts[0].items i ON TRUE
    LEFT JOIN i.item_totals it ON TRUE
    LEFT JOIN it.details d ON TRUE
    GROUP BY b.ref, TRIM('"' FROM CAST(i.sku AS VARCHAR))
)
SELECT
     oi.ref
    ,oi.sku
    ,md.categoria
    ,oi.name
    ,bso.custom_manifest_number AS manifiesto
    ,oi.quantity
    ,md.quantity_real
    ,oi.item_cost
    ,oi.fob_item
    ,SUM(oi.fob_item) OVER (PARTITION BY oi.ref) AS fob_orden
    ,md.total_weight
    ,tx.tax_total
FROM order_items AS oi
LEFT JOIN manifest_data AS md
    ON  oi.ref       = md.ref
    AND oi.clean_sku = md.clean_sku
    AND md.rn        = 1
LEFT JOIN sales_orders AS bso
    ON oi.ref = bso.ref_id
LEFT JOIN tax_income_item AS tx
    ON  oi.ref       = tx.ref
    AND oi.clean_sku = UPPER(REPLACE(SUBSTRING(tx.sku, CHARINDEX('-', tx.sku) + 1), '_', ''))
ORDER BY oi.ref;
            """
            print(f"[taxes_thread] Ejecutando query Redshift para {len(ordenes)} órdenes...")
            log_text.insert('end', f"Ejecutando query Redshift para {len(ordenes)} órdenes...\n")
            cur1.execute(query)
            rows = cur1.fetchall()
            colnames = [desc[0] for desc in cur1.description]
            cur1.close()
            conn1.close()
            print(f"[taxes_thread] Resultados obtenidos: {len(rows)} filas")
        except Exception as e:
            print(f"[taxes_thread] Error consultando Redshift: {e}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', f"Error consultando Redshift: {e}\n")
            return


        # --- NUEVO: Limpiar y pegar resultados en hoja INPUT TM ---
        col_end_clear = chr(ord('A') + len(colnames) - 1)
        clear_range_tm = f"{SHEET_TM}!A2:{col_end_clear}"
        print(f"[taxes_thread] Limpiando hoja {SHEET_TM} rango {clear_range_tm}")
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=SHEET_ID,
            range=clear_range_tm
        ).execute()

        # Helper para convertir Decimals y otros tipos a formatos compatibles
        def convert_cell(cell):
            if isinstance(cell, decimal.Decimal):
                return float(cell)
            if isinstance(cell, bytes):
                return cell.decode('utf-8')
            return cell

        # Preparar los datos para pegar (sin encabezados)
        values_tm = [[convert_cell(cell) for cell in row] for row in rows]
        start_row_tm = 2
        end_row_tm = start_row_tm + len(values_tm) - 1
        col_end_tm = chr(ord('A') + len(colnames) - 1)
        range_destino_tm = f"{SHEET_TM}!A{start_row_tm}:{col_end_tm}{end_row_tm}"
        print(f"[taxes_thread] Pegando resultados en {range_destino_tm}")
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=range_destino_tm,
            valueInputOption="USER_ENTERED",  # Permite que Google Sheets interprete números automáticamente
            body={"values": values_tm}
        ).execute()
        print(f"[taxes_thread] Resultados pegados en hoja {SHEET_TM}.")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"Resultados de Redshift pegados en hoja {SHEET_TM}.\n")

        # --- Clasificación / Courier5 tax_total replacement ---
        print("[taxes_thread] Iniciando paso de clasificación / extracción de taxes...")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', "Clasificando / extrayendo taxes por SKU...\n")

        try:
            # Primero limpiar la columna de clasificación (columna después de los colnames)
            col_clasificacion = chr(ord('A') + len(colnames))
            clear_range_clasificacion = f"{SHEET_TM}!{col_clasificacion}2:{col_clasificacion}"
            print(f"[taxes_thread] Limpiando columna de clasificaciones: {clear_range_clasificacion}")
            sheets_service.spreadsheets().values().clear(
                spreadsheetId=SHEET_ID,
                range=clear_range_clasificacion
            ).execute()

            # Si es Courier5 (Uruguay) no usamos la IA: consultamos Redshift por tax_total por increment_id y sku
            if courier == 'Courier5' or pais and pais.lower().startswith('urug'):
                print("[taxes_thread] Courier Courier5 detectado: obteniendo tax_total desde Redshift en lugar de IA.")
                try:
                    conn_Courier5 = psycopg2.connect(
                        host=os.getenv('REDSHIFT_HOST'),
                        port=os.getenv('REDSHIFT_PORT'),
                        dbname=os.getenv('REDSHIFT_DB'),
                        user=usuario_redshift,
                        password=contrasena_redshift
                    )
                    cur_Courier5 = conn_Courier5.cursor()

                    # Armar lista de órdenes para la query
                    ordenes_sql = ','.join([f"'{o}'" for o in ordenes])

                    query_Courier5 = f"""
WITH base_data AS (
    SELECT 
        increment_id,
        pricing_data AS p_data 
    FROM tm_staging.m2_sales_order_delsert
    WHERE increment_id IN ({ordenes_sql})
)
SELECT 
    b.increment_id,
    TRIM('"' FROM CAST(i.sku AS VARCHAR)) AS sku,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'subtotal' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS subtotal,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'customfee' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS custom_fee,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'customfee_discount' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS discount,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'tax' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS tax_total,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'shipping' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS shipping,
    MAX(CASE WHEN TRIM('"' FROM CAST(d.code AS VARCHAR)) = 'uy_tax_import_plane' THEN CAST(d.amount AS DECIMAL(10,2)) END) AS tax_import_plane,
    MAX(CASE WHEN TRIM('"' FROM CAST(d.code AS VARCHAR)) LIKE 'uy_iva%%' THEN CAST(d.amount AS DECIMAL(10,2)) END) AS tax_iva
FROM base_data b
LEFT JOIN b.p_data.carts[0].items i ON TRUE
LEFT JOIN i.item_totals it ON TRUE
LEFT JOIN it.details d ON TRUE
GROUP BY b.increment_id, TRIM('"' FROM CAST(i.sku AS VARCHAR));
                    """

                    print(f"[taxes_thread] Ejecutando query Courier5 tax_total para {len(ordenes)} órdenes...")
                    cur_Courier5.execute(query_Courier5)
                    Courier5_rows = cur_Courier5.fetchall()
                    Courier5_cols = [d[0] for d in cur_Courier5.description]

                    # Mapear (increment_id, sku) -> tax_total
                    idx_inc = Courier5_cols.index('increment_id') if 'increment_id' in Courier5_cols else 0
                    idx_sku = Courier5_cols.index('sku') if 'sku' in Courier5_cols else 1
                    idx_tax = Courier5_cols.index('tax_total') if 'tax_total' in Courier5_cols else None

                    tax_map = {}
                    for ur in Courier5_rows:
                        key = (str(ur[idx_inc]).strip(), str(ur[idx_sku]).strip())
                        tax_map[key] = ur[idx_tax] if idx_tax is not None else None

                    # Construir lista de valores alineada con 'rows' (resultados del query principal)
                    ref_idx = colnames.index('ref') if 'ref' in colnames else 0
                    sku_idx = colnames.index('sku') if 'sku' in colnames else 1

                    clasificaciones = []
                    for r in rows:
                        ref = str(r[ref_idx]).strip()
                        sku = str(r[sku_idx]).strip()
                        tax_val = tax_map.get((ref, sku), '')
                        # Normalizar Decimal a float si aplica
                        if isinstance(tax_val, decimal.Decimal):
                            tax_val = float(tax_val)
                        clasificaciones.append(tax_val)

                    # Pegar clasificaciones (tax_total) en columna de clasificación
                    if clasificaciones:
                        range_clasificacion = f"{SHEET_TM}!{col_clasificacion}{start_row_tm}:{col_clasificacion}{end_row_tm}"
                        print(f"[taxes_thread] Pegando tax_total en {range_clasificacion}")
                        sheets_service.spreadsheets().values().update(
                            spreadsheetId=SHEET_ID,
                            range=range_clasificacion,
                            valueInputOption="USER_ENTERED",
                            body={"values": [[c] for c in clasificaciones]}
                        ).execute()
                        if 'log_text' in globals() and log_text:
                            log_text.insert('end', f"✅ {len(clasificaciones)} tax_total pegados para Courier5.\n")

                    cur_Courier5.close()
                    conn_Courier5.close()
                except Exception as e_Courier5:
                    print(f"[taxes_thread] Error consultando tax_total Courier5: {e_Courier5}")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"⚠️ Error consultando tax_total Courier5: {e_Courier5}\n")
            else:
                # Flujo normal: clasificación por IA
                clasificaciones = clasificar_productos_con_ia(
                    rows,
                    pais,
                    courier,
                    sheets_service=sheets_service,
                    sheet_id=SHEET_ID,
                    sheet_name=SHEET_TM,
                    start_row=start_row_tm,
                    col=col_clasificacion
                )

                # Pegar clasificaciones finales en columna de clasificación (después de los datos)
                if clasificaciones:
                    range_clasificacion = f"{SHEET_TM}!{col_clasificacion}{start_row_tm}:{col_clasificacion}{end_row_tm}"
                    print(f"[taxes_thread] Pegando clasificaciones finales en {range_clasificacion}")
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SHEET_ID,
                        range=range_clasificacion,
                        valueInputOption="USER_ENTERED",  # Permite que Google Sheets interprete valores automáticamente
                        body={"values": [[c] for c in clasificaciones]}
                    ).execute()
                    print(f"[taxes_thread] Clasificaciones pegadas exitosamente.")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"✅ {len(clasificaciones)} productos clasificados.\n")

            # Consolidar INPUT TM - Pos al historial por país
            _taxes_consolidar_input_tm_a_historial(
                sheets_service,
                SHEET_ID,
                SHEET_SOURCE,
                courier,
                log_text,
                key_cols=(0, 1),
                tipo="Pos",
            )
            # Consolidar INPUT Courier al historial centralizado
            _taxes_consolidar_input_tm_a_historial(
                sheets_service,
                SHEET_ID,
                SHEET_NAME,
                courier,
                log_text,
                key_cols=(0,),
                tipo=None,
                destination_sheet_name_override="Consolidado Input Courier",
            )
        except Exception as e:
            print(f"[taxes_thread] Error en clasificación/extracción de taxes: {e}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', f"⚠️ Error en clasificación/extracción: {e}\n")

    except Exception as e:
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"Error en taxes_thread: {e}\n")

        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )

    finally:
        print("[taxes_thread] Proceso finalizado.")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', "Finalizado\n")

        registrar_accion(
            total_orders=len(ordenes),  
            agente=agente_inconv_var.get(),
        )

def pretaxes_thread():
    """Pre-taxes: obtiene órdenes a partir de un número de manifiesto,
    corre la query principal de Redshift, clasifica con IA y consolida
    en la hoja INPUT TM - Pre y luego en el consolidado por país con tipo='Pre'.
    No toca INPUT Courier.
    """
    global select_option_var, usuario_redshift, contrasena_redshift, usuario_redshift_var, \
        contrasena_redshift_var, seller, courier_var, manifesto_var, action_var, \
        username_company_entry, password_company_entry, usuario_company, \
        contrasena_company, username_entry, log_text, usuario, contrasena, \
        frame_left, frame_right_2, frame_right, frame_middle, password_entry, \
        usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, \
        tbody_elements_arrived, sku_find

    import datetime
    import os
    import decimal
    import psycopg2
    from dotenv import load_dotenv
    load_dotenv()

    print(usuario_redshift, contrasena_redshift)

    # Validar conexión a Redshift antes de todo
    try:
        conn_test = psycopg2.connect(
            host=os.getenv('REDSHIFT_HOST'),
            port=os.getenv('REDSHIFT_PORT'),
            dbname=os.getenv('REDSHIFT_DB'),
            user=usuario_redshift,
            password=contrasena_redshift,
            connect_timeout=5
        )
        conn_test.close()
        print("[pretaxes_thread] Redshift credentials OK")
    except Exception as auth_err:
        msg = f"⚠️ Error conectando a Redshift: {auth_err}"
        print(msg)
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    courier = courier_var.get() if 'courier_var' in globals() else None
    manifiesto = manifesto_var.get().strip() if 'manifesto_var' in globals() and manifesto_var else ''

    if not courier or courier.strip() == "":
        msg = "⚠️ Debe seleccionar un courier."
        print(f"[pretaxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    courier = courier.strip().upper()

    if not manifiesto:
        msg = "⚠️ Ingresá el número de manifiesto antes de ejecutar."
        print(f"[pretaxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    # Configuración por courier (mismos SHEET_ID que taxes)
    if courier == "Courier1":
        SHEET_ID = "1D0YFVnWzbsnmi5gZNpJ6XVENgC5Blo116OpQ-evggzc"
        SHEET_TM = "INPUT TM - Pre"
        pais = "Argentina"
    elif courier == "Courier2":
        SHEET_ID = "1oy0bXnZGzknb91jFwXmtc_QGx0tyzY7AcVc19iRUHNc"
        SHEET_TM = "INPUT TM - Pre"
        pais = "Costa Rica"
    elif courier == "Courier3":
        SHEET_ID = "1QNsX-_j-_hdiGkf79czaLCkePX4x5n3rmN0bxK7uu7I"
        SHEET_TM = "INPUT TM - Pre"
        pais = "Peru"
    elif courier == "Courier4":
        SHEET_ID = "1FW71FZ2lBktZuUOSjWQnOChLpA7JCzBhmDv9uqLRhMM"
        SHEET_TM = "INPUT TM - Pre"
        pais = "Ecuador"
    elif courier == "Courier5":
        SHEET_ID = "1b8tTL-0Z7DZEER-ewtNp4O7NpWZLqchTHFmLu7sT0eE"
        SHEET_TM = "INPUT TM - Pre"
        pais = "Uruguay"
    else:
        msg = f"⚠️ Courier no soportado: {courier}"
        print(f"[pretaxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    CRED_FILE = "credenciales_drive.json"

    import io
    from googleapiclient.discovery import build
    from google.oauth2.service_account import Credentials

    ordenes = []
    try:
        creds = Credentials.from_service_account_file(
            CRED_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"]
        )
        sheets_service = build("sheets", "v4", credentials=creds)
    except Exception as e:
        msg = f"⚠️ Error de autenticación Google: {e}"
        print(f"[pretaxes_thread] {msg}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', msg + '\n')
        return

    try:
        # Buscar órdenes por manifiesto en Redshift
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"🔍 Buscando órdenes para manifiesto '{manifiesto}'...\n")
        print(f"[pretaxes_thread] Buscando órdenes para manifiesto: {manifiesto}")

        conn_mani = psycopg2.connect(
            host=os.getenv('REDSHIFT_HOST'),
            port=os.getenv('REDSHIFT_PORT'),
            dbname=os.getenv('REDSHIFT_DB'),
            user=usuario_redshift,
            password=contrasena_redshift
        )
        cur_mani = conn_mani.cursor()
        pattern = f"{manifiesto}%"
        query_mani = """
            SELECT DISTINCT ref_id
            FROM tm_staging.brightpearl_sales_orders
            WHERE ref_id IS NOT NULL
              AND UPPER(custom_manifest_number) LIKE UPPER(%s)
        """
        cur_mani.execute(query_mani, (pattern,))
        mani_rows = cur_mani.fetchall()
        cur_mani.close()
        conn_mani.close()

        ordenes = [str(r[0]).strip() for r in mani_rows if r[0]]
        print(f"[pretaxes_thread] Órdenes encontradas: {len(ordenes)}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"✅ {len(ordenes)} órdenes encontradas para el manifiesto.\n")

        if not ordenes:
            msg = f"⚠️ No se encontraron órdenes para el manifiesto '{manifiesto}'."
            print(f"[pretaxes_thread] {msg}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', msg + '\n')
            return

        # Query principal de Redshift (idéntica a taxes)
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"Ejecutando query Redshift para {len(ordenes)} órdenes...\n")
        ordenes_sql = ','.join([f"'{o}'" for o in ordenes])
        query = f"""
WITH order_items AS (
    SELECT
         omsoi.ref_id         AS ref
        ,omsoi.item_sku       AS sku
        ,UPPER(REPLACE(SUBSTRING(omsoi.item_sku, CHARINDEX('-', omsoi.item_sku) + 1), '_', '')) AS clean_sku
        ,omsoi.item_cost
        ,omsoi.item_net       AS fob_item
        ,omsoi.totalvalue_net AS fob_orden
        ,omsoi.item_name      AS name
        ,omsoi.quantity
        ,omsoi.updated_at_oms
    FROM tm_staging.brightpearl_purchase_order_item AS omsoi
    WHERE omsoi.ref_id IN ({ordenes_sql})
        AND omsoi.item_sku IS NOT NULL
        AND omsoi.item_sku <> ''
        AND omsoi.statusid_oms IN ('1.1 Placed with supplier', '3.1 Invoice received')
), sales_orders AS (
    SELECT DISTINCT
         ref_id
        ,custom_manifest_number
    FROM tm_staging.brightpearl_sales_orders
    WHERE ref_id IN (SELECT ref FROM order_items)
), manifest_data AS (
    SELECT
         mdi.manifest_data_order_mgt AS ref
        ,mdi.sku                     AS sku
        ,UPPER(REPLACE(REPLACE(SUBSTRING(mdi.sku, CHARINDEX('-', mdi.sku) + 1), '_', ''), '|', '')) AS clean_sku
        ,mdi.category                AS categoria
        ,mdi.total_weight
        ,mdi.quantity_real
        ,mdi.price
        ,ROW_NUMBER() OVER (
            PARTITION BY mdi.manifest_data_order_mgt, mdi.sku
            ORDER BY mdi.manifest_data_order_mgt DESC
        ) AS rn
    FROM tm_staging.m1_manifest_data_items_delsert AS mdi
    INNER JOIN order_items AS oi
        ON mdi.manifest_data_order_mgt = oi.ref
    WHERE mdi.url <> 'non-image'
), tax_raw AS (
    SELECT increment_id AS ref, pricing_data AS p_data
    FROM tm_staging.m2_sales_order_delsert
    WHERE increment_id IN (SELECT ref FROM order_items)
), tax_income_item AS (
    SELECT
         b.ref
        ,TRIM('"' FROM CAST(i.sku AS VARCHAR)) AS sku
        ,MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'tax' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS tax_total
    FROM tax_raw AS b
    LEFT JOIN b.p_data.carts[0].items i ON TRUE
    LEFT JOIN i.item_totals it ON TRUE
    LEFT JOIN it.details d ON TRUE
    GROUP BY b.ref, TRIM('"' FROM CAST(i.sku AS VARCHAR))
)
SELECT
     oi.ref
    ,oi.sku
    ,md.categoria
    ,oi.name
    ,bso.custom_manifest_number AS manifiesto
    ,oi.quantity
    ,md.quantity_real
    ,oi.item_cost
    ,oi.fob_item
    ,SUM(oi.fob_item) OVER (PARTITION BY oi.ref) AS fob_orden
    ,md.total_weight
    ,tx.tax_total
FROM order_items AS oi
LEFT JOIN manifest_data AS md
    ON  oi.ref       = md.ref
    AND oi.clean_sku = md.clean_sku
    AND md.rn        = 1
LEFT JOIN sales_orders AS bso
    ON oi.ref = bso.ref_id
LEFT JOIN tax_income_item AS tx
    ON  oi.ref       = tx.ref
    AND oi.clean_sku = UPPER(REPLACE(SUBSTRING(tx.sku, CHARINDEX('-', tx.sku) + 1), '_', ''))
ORDER BY oi.ref;
        """
        conn2 = psycopg2.connect(
            host=os.getenv('REDSHIFT_HOST'),
            port=os.getenv('REDSHIFT_PORT'),
            dbname=os.getenv('REDSHIFT_DB'),
            user=usuario_redshift,
            password=contrasena_redshift
        )
        cur2 = conn2.cursor()
        cur2.execute(query)
        rows = cur2.fetchall()
        colnames = [desc[0] for desc in cur2.description]
        cur2.close()
        conn2.close()
        print(f"[pretaxes_thread] Resultados obtenidos: {len(rows)} filas")

        # Limpiar y pegar en INPUT TM - Pre
        def convert_cell(cell):
            if isinstance(cell, decimal.Decimal):
                return float(cell)
            if isinstance(cell, bytes):
                return cell.decode('utf-8')
            return cell

        col_end_clear = chr(ord('A') + len(colnames) - 1)
        clear_range_tm = f"{SHEET_TM}!A2:{col_end_clear}"
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=SHEET_ID,
            range=clear_range_tm
        ).execute()

        values_tm = [[convert_cell(cell) for cell in row] for row in rows]
        start_row_tm = 2
        end_row_tm = start_row_tm + len(values_tm) - 1
        col_end_tm = chr(ord('A') + len(colnames) - 1)
        range_destino_tm = f"{SHEET_TM}!A{start_row_tm}:{col_end_tm}{end_row_tm}"
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=range_destino_tm,
            valueInputOption="USER_ENTERED",
            body={"values": values_tm}
        ).execute()
        print(f"[pretaxes_thread] Resultados pegados en hoja {SHEET_TM}.")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"Resultados de Redshift pegados en hoja {SHEET_TM}.\n")

        # Clasificación / extracción de taxes
        try:
            col_clasificacion = chr(ord('A') + len(colnames))
            clear_range_clas = f"{SHEET_TM}!{col_clasificacion}2:{col_clasificacion}"
            sheets_service.spreadsheets().values().clear(
                spreadsheetId=SHEET_ID,
                range=clear_range_clas
            ).execute()

            if courier == 'Courier5' or (pais and pais.lower().startswith('urug')):
                # Courier5: usar tax_total desde Redshift
                conn_Courier5 = psycopg2.connect(
                    host=os.getenv('REDSHIFT_HOST'),
                    port=os.getenv('REDSHIFT_PORT'),
                    dbname=os.getenv('REDSHIFT_DB'),
                    user=usuario_redshift,
                    password=contrasena_redshift
                )
                cur_Courier5 = conn_Courier5.cursor()
                query_Courier5 = f"""
WITH base_data AS (
    SELECT increment_id, pricing_data AS p_data
    FROM tm_staging.m2_sales_order_delsert
    WHERE increment_id IN ({ordenes_sql})
)
SELECT
    b.increment_id,
    TRIM('"' FROM CAST(i.sku AS VARCHAR)) AS sku,
    MAX(CASE WHEN TRIM('"' FROM CAST(it.code AS VARCHAR)) = 'tax' THEN CAST(it.amount AS DECIMAL(10,2)) END) AS tax_total
FROM base_data b
LEFT JOIN b.p_data.carts[0].items i ON TRUE
LEFT JOIN i.item_totals it ON TRUE
LEFT JOIN it.details d ON TRUE
GROUP BY b.increment_id, TRIM('"' FROM CAST(i.sku AS VARCHAR));
                """
                cur_Courier5.execute(query_Courier5)
                Courier5_rows = cur_Courier5.fetchall()
                Courier5_cols = [d[0] for d in cur_Courier5.description]
                cur_Courier5.close()
                conn_Courier5.close()

                idx_inc = Courier5_cols.index('increment_id') if 'increment_id' in Courier5_cols else 0
                idx_sku_u = Courier5_cols.index('sku') if 'sku' in Courier5_cols else 1
                idx_tax = Courier5_cols.index('tax_total') if 'tax_total' in Courier5_cols else None
                tax_map = {}
                for ur in Courier5_rows:
                    key = (str(ur[idx_inc]).strip(), str(ur[idx_sku_u]).strip())
                    tax_map[key] = ur[idx_tax] if idx_tax is not None else None

                ref_idx = colnames.index('ref') if 'ref' in colnames else 0
                sku_idx = colnames.index('sku') if 'sku' in colnames else 1
                clasificaciones = []
                for r in rows:
                    ref = str(r[ref_idx]).strip()
                    sku = str(r[sku_idx]).strip()
                    tax_val = tax_map.get((ref, sku), '')
                    if isinstance(tax_val, decimal.Decimal):
                        tax_val = float(tax_val)
                    clasificaciones.append(tax_val)

                if clasificaciones:
                    range_clas = f"{SHEET_TM}!{col_clasificacion}{start_row_tm}:{col_clasificacion}{end_row_tm}"
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SHEET_ID,
                        range=range_clas,
                        valueInputOption="USER_ENTERED",
                        body={"values": [[c] for c in clasificaciones]}
                    ).execute()
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"✅ {len(clasificaciones)} tax_total pegados para Courier5.\n")
            else:
                # Flujo normal: clasificación por IA
                clasificaciones = clasificar_productos_con_ia(
                    rows,
                    pais,
                    courier,
                    sheets_service=sheets_service,
                    sheet_id=SHEET_ID,
                    sheet_name=SHEET_TM,
                    start_row=start_row_tm,
                    col=col_clasificacion
                )
                if clasificaciones:
                    range_clas = f"{SHEET_TM}!{col_clasificacion}{start_row_tm}:{col_clasificacion}{end_row_tm}"
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SHEET_ID,
                        range=range_clas,
                        valueInputOption="USER_ENTERED",
                        body={"values": [[c] for c in clasificaciones]}
                    ).execute()
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"✅ {len(clasificaciones)} productos clasificados.\n")

            # Consolidar INPUT TM - Pre al historial por país con tipo='Pre'
            _taxes_consolidar_input_tm_a_historial(
                sheets_service,
                SHEET_ID,
                SHEET_TM,
                courier,
                log_text,
                key_cols=(0, 1),
                tipo="Pre",
            )
        except Exception as e:
            print(f"[pretaxes_thread] Error en clasificación: {e}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', f"⚠️ Error en clasificación: {e}\n")

    except Exception as e:
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"Error en pretaxes_thread: {e}\n")
        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get() if 'agente_inconv_var' in globals() else 'N/A',
            order_id="Error general",
            exception=e
        )

    finally:
        print("[pretaxes_thread] Proceso finalizado.")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', "Finalizado\n")
        registrar_accion(
            total_orders=len(ordenes),
            agente=agente_inconv_var.get() if 'agente_inconv_var' in globals() else 'N/A',
        )


def extraer_nombre_pdf():
    # --- Configuración ---
    SPREADSHEET_ID = "1WjoEORNcH_S54N2o-zbRj1JrTZgM4hzFAKNIGjUVxS8"
    NOMBRE_HOJA = "Historial"
    COLUMNA_B = 1
    COLUMNA_D = 3
    COLUMNA_RESULTADO = 5  # columna F (índice base 0)
    CRED_FILE = "credenciales_drive.json"

    print("🔑 Autenticando con Google...")
    if 'log_text' in globals() and log_text:
        log_text.insert('end', "🔑 Autenticando con Google...\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive.readonly"]
    credentials = Credentials.from_service_account_file(CRED_FILE, scopes=scopes)

    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(NOMBRE_HOJA)
    drive_service = build("drive", "v3", credentials=credentials)

    print("📄 Leyendo datos de la hoja...")
    if 'log_text' in globals() and log_text:
        log_text.insert('end', "📄 Leyendo datos de la hoja...\n")
    data = sheet.get_all_values()

    cell_list = []
    print(f"📊 Total de filas en la hoja: {len(data)}")

    for i in range(1, len(data)):  # saltar encabezado
        fila = data[i]
        ref = fila[COLUMNA_B] if len(fila) > COLUMNA_B else ""
        url = fila[COLUMNA_D] if len(fila) > COLUMNA_D else ""
        resultado_actual = fila[COLUMNA_RESULTADO] if len(fila) > COLUMNA_RESULTADO else ""

        # Solo procesar si la columna F está vacía
        if resultado_actual:
            continue

        print(f"🔍 Procesando fila {i+1}: ref={ref}, url={url}")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"🔍 Procesando fila {i+1}: ref={ref}, url={url}\n")

        if not url:
            resultado = "Sin link"
            print(f"⚠️ Sin link en fila {i+1}")
            if 'log_text' in globals() and log_text:
                log_text.insert('end', f"⚠️ Sin link en fila {i+1}\n")
        else:
            match = re.search(r"/d/([a-zA-Z0-9_-]{25,})", url)
            if not match:
                resultado = "Link inválido"
                print(f"❌ Link inválido en fila {i+1}")
                if 'log_text' in globals() and log_text:
                    log_text.insert('end', f"❌ Link inválido en fila {i+1}\n")
            else:
                file_id = match.group(1)
                try:
                    file = drive_service.files().get(
                        fileId=file_id,
                        fields="name",
                        supportsAllDrives=True
                    ).execute()
                    nombre = file["name"].replace(".pdf", "").strip()
                    resultado = str(nombre).lower() == str(ref).strip().lower()
                    resultado = str(resultado).upper()
                    print(f"✅ Fila {i+1}: nombre en Drive='{nombre}', ref='{ref}', resultado={resultado}")
                    if 'log_text' in globals() and log_text:
                        log_text.insert('end', f"✅ Fila {i+1}: nombre en Drive='{nombre}', ref='{ref}', resultado={resultado}\n")
                except Exception as e:
                    if "File not found" in str(e):
                        resultado = "Archivo no encontrado"
                        print(f"⚠️ Archivo no encontrado en fila {i+1}")
                        if 'log_text' in globals() and log_text:
                            log_text.insert('end', f"⚠️ Archivo no encontrado en fila {i+1}\n")
                    else:
                        resultado = f"Error: {str(e)}"
                        print(f"⚠️ Error en fila {i+1}: {str(e)}")
                        if 'log_text' in globals() and log_text:
                            log_text.insert('end', f"⚠️ Error en fila {i+1}: {str(e)}\n")

        # Solo agregar si la fila tiene algún dato relevante (por ejemplo, ref o url)
        if ref or url:
            cell = Cell(row=i + 1, col=COLUMNA_RESULTADO + 1, value=resultado)
            cell_list.append(cell)

    if cell_list:
        print(f"⬆️ Actualizando hoja en {len(cell_list)} celdas de la columna F...")
        if 'log_text' in globals() and log_text:
            log_text.insert('end', f"⬆️ Actualizando hoja en {len(cell_list)} celdas de la columna F...\n")
        sheet.update_cells(cell_list)

    print(f"✅ {len(cell_list)} filas procesadas.")
    if 'log_text' in globals() and log_text:
        log_text.insert('end', f"✅ {len(cell_list)} filas procesadas.\n")

def extraer_pdf():
    # 📁 Ruta a la carpeta de descargas
    home_dir = os.path.expanduser("~")
    downloads_folder_en = os.path.join(home_dir, "Downloads")
    downloads_folder_es = os.path.join(home_dir, "Descargas")

    if os.path.isdir(downloads_folder_en):
        downloads_folder = downloads_folder_en
    elif os.path.isdir(downloads_folder_es):
        downloads_folder = downloads_folder_es
    else:
        downloads_folder = downloads_folder

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

    # 🔍 Obtener solo archivos PDF
    pdf_files = [f for f in os.listdir(downloads_folder) if f.lower().endswith('.pdf')]

    # 📄 Abrir hoja existente (reemplaza con el nombre correcto)
    spreadsheet = cliente.open("Bot changes Courier6 ")  # <-- Modificá esto
    hoja = spreadsheet.worksheet("Nombres de pdf")  # o .worksheet("Nombre de la pestaña")

    # 🧾 Crear los datos a subir
    data = [["Nombre del PDF"]] + [[nombre] for nombre in pdf_files]

    # ⬆️ Subir de una sola vez
    hoja.update(range_name="A1", values=data)

    print("✅ Listado de PDFs actualizado con éxito.")

    log_text.insert(tk.END, f"Finalizado\n")

def asignar_track_MA():
    global select_option_var, usuario_db_entry, password_db_entry, usuario_db, contrasena_db, seller, action_var, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, usuario, contrasena, valores_coma_separada, cuenta_version, tbody_elements, tbody_elements_arrived, sku_find

    selected_column = action_var.get()
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
    agente = agente_inconv_var.get()
    print("Agente seleccionado:", agente)
    idx_general = 0  
    ordenes_procesadas = 0 
    spreadsheet_id = "1V5QCYj0x3FiTGHCmlp2wpVN_sOPKWrAKsmcNuDAy3Ac"
    historial = cliente.open_by_key(spreadsheet_id).worksheet("Historial")
    data = historial.get_all_values()
    try:
        # Configuración de Selenium
        chrome_options = Options()
        selenium_service = None
        try:
            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(
                service=Service(driver_path),
                options=chrome_options
            )
        except Exception:
            driver = webdriver.Chrome(
                service=selenium_service,
                options=chrome_options
            )
        set_company_custom_header(driver)

        # Página web oms
        driver.get("https://use1.omsapp.com/admin_login.php?clients_id=company")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email_address']")))
        username.clear()
        username.send_keys(usuario)
        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password.clear()
        password.send_keys(contrasena)
        button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, 'submit-admin'))).click()

        time.sleep(5)

        # Página web mail Courier8
        driver.get("https://shipping.Courier8.com/login")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 2).until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.stop();")

        # Inicio de sesión
        username = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email']")))
        username.clear()
        username.send_keys(usuario_company)
        password = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        password.clear()
        password.send_keys(contrasena_company)
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary.btn-lg.block.full-width.m-t")))
        btn.click()

        for idx, fila in enumerate(data[1:], start=2):  # Empieza en la fila 2 (1-based)
            so = fila[0] if len(fila) > 0 else ""
            ref = fila[1] if len(fila) > 1 else ""
            tracking = fila[7] if len(fila) > 7 else ""
            
            # VALIDAR QUE SO Y REF NO ESTÉN VACÍOS
            if not so or not so.strip():
                log_text.insert(tk.END, f"⚠️ Fila {idx} sin SO, saltando...\n")
                print(f"⚠️ Fila {idx} sin SO, saltando...")
                continue
                
            if not ref or not ref.strip():
                log_text.insert(tk.END, f"⚠️ Fila {idx} sin REF, saltando...\n")
                print(f"⚠️ Fila {idx} sin REF, saltando...")
                continue
            
            if not tracking.strip():
                ordenes_procesadas += 1  # <-- Incrementar contador
                log_text.insert(tk.END, f"Procesando orden: {ref} - {so} (fila {idx})\n")
                print(f"Procesando orden: {ref} - {so} (fila {idx})")
                file_link = []
                files_link = []

                url = f"https://shipping.Courier8.com/packages?created_from=&created_to=&tracking_code=&order_id={ref}"
                driver.get(url)
                time.sleep(3)
                tracking = None

                try:
                    # Encuentra la tabla
                    tabla = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "table.table.table-hover.table-stripped.table-condensed.packages-table"))
                    )

                    # Encuentra todas las filas del cuerpo de la tabla
                    filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

                    # Itera sobre las filas para encontrar el valor de la columna "Tracking Number"

                    for fila in filas:
                        columnas = fila.find_elements(By.TAG_NAME, "td")
                        if len(columnas) > 6:  # Asegúrate de que la columna "Tracking Number" existe
                            tracking = columnas[6].text  # Columna 7 (índice 6)

                    if not tracking or not re.match(r"^[A-Z]{2}\d{9}[A-Z]{2}$", tracking):
                        # Si no cumple, busca en el texto de la página
                        page_text = driver.find_element(By.TAG_NAME, "body").text
                        match = re.search(r"[A-Z]{2}\d{9}[A-Z]{2}", page_text)
                        if match:
                            tracking = match.group(0)
                            print("Tracking encontrado en el texto de la página:", tracking)
                        else:
                            print("No se encontró tracking válido en la tabla ni en el texto de la página.")
                    else:
                        print("Tracking encontrado:", tracking)
                except Exception as e:
                    try:
                        print(f"No se pudo obtener el tracking en mail Courier8, segundo intento {ref}: {str(e)}")
                        url = f"https://shipping.Courier8.com/packages?created_from=&created_to=&tracking_code=&order_id={ref}"
                        driver.get(url)
                        time.sleep(5)

                        # Encuentra la tabla
                        tabla = WebDriverWait(driver, 40).until(
                            EC.presence_of_element_located(
                                (By.CSS_SELECTOR,
                                 "table.table.table-hover.table-stripped.table-condensed.packages-table"))
                        )

                        # Encuentra todas las filas del cuerpo de la tabla
                        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

                        # Itera sobre las filas para encontrar el valor de la columna "Tracking Number"
                        for fila in filas:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            if len(columnas) > 6:  # Asegúrate de que la columna "Tracking Number" existe
                                tracking = columnas[6].text  # Columna 7 (índice 6)

                        if not tracking or not re.match(r"^[A-Z]{2}\d{9}[A-Z]{2}$", tracking):
                            # Si no cumple, busca en el texto de la página
                            page_text = driver.find_element(By.TAG_NAME, "body").text
                            match = re.search(r"[A-Z]{2}\d{9}[A-Z]{2}", page_text)
                            if match:
                                tracking = match.group(0)
                                print("Tracking encontrado en el texto de la página:", tracking)
                            else:
                                print("No se encontró tracking válido en la tabla ni en el texto de la página.")
                        else:
                            print("Tracking encontrado:", tracking)

                    except Exception as e:
                        print(f"No se pudo obtener el tracking en mail Courier8 {ref}: {str(e)}")
                        log_text.insert(tk.END, f"No se pudo obtener el tracking en mail Courier8 {ref}\n")

                        if tracking:
                            historial.update_cell(idx, 8, tracking)
                        continue

                if tracking:
                    try:
                        time.sleep(2)
                        # Entrar a PO y accionar
                        invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                        driver.get(invoice_url)

                        # Obtener ref
                        elemento_id = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, 'orders_customer_ref'))).get_attribute('value')
                        ref_oms = elemento_id
                        print(ref_oms)

                        if not ref == ref_oms:
                            print(f"La ref {ref} no coincide con la ref en oms {ref_oms}")
                            log_text.insert(tk.END, f"La ref {ref} no coincide con la ref en oms {ref_oms}\n")
                            if tracking:
                                historial.update_cell(idx, 8, tracking)
                            continue

                        button_Custom = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                        button_Custom.click()

                        xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                        try:
                            el_shipping = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, xpath_shipping))
                            )

                            # Traer a vista
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el_shipping)
                            time.sleep(0.2)

                            # Intento normal esperando que sea clickeable
                            try:
                                WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                el_shipping.click()
                            except (ElementNotInteractableException, ElementClickInterceptedException,
                                    TimeoutException):
                                # Fallback 1: click vía JS
                                try:
                                    driver.execute_script("arguments[0].click();", el_shipping)
                                except Exception:
                                    # Fallback 2: mover con Actions y click
                                    try:
                                        ActionChains(driver).move_to_element(el_shipping).click().perform()
                                    except Exception:
                                        # Fallback 3: focus y enviar ENTER
                                        try:
                                            driver.execute_script("arguments[0].focus();", el_shipping)
                                            el_shipping.send_keys(Keys.ENTER)
                                        except Exception as final_err:
                                            print("No se pudo hacer click en Shipping:", final_err)

                        except TimeoutException:
                            print("No se encontró el tab Shipping en la página.")

                        # Localiza el campo de entrada por su ID
                        input_field = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                        )

                        # Asegurar visibilidad antes de click
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                        time.sleep(0.3)
                        try:
                            input_field.clear()

                            input_field.send_keys(tracking)
                        except Exception:
                            # fallback por si el click normal falla
                            driver.execute_script(
                                "arguments[0].value = arguments[1];"
                                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                                input_field, tracking
                            )

                        # Guardar progreso
                        button_save_clon = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                        try:
                            time.sleep(5)
                            # Entrar a PO y accionar
                            invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                            driver.get(invoice_url)

                            button_Custom = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                            button_Custom.click()

                            xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                            try:
                                el_shipping = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located((By.XPATH, xpath_shipping))
                                )

                                # Traer a vista
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el_shipping)
                                time.sleep(0.2)

                                # Intento normal esperando que sea clickeable
                                try:
                                    WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                    el_shipping.click()
                                except (ElementNotInteractableException, ElementClickInterceptedException,
                                        TimeoutException):
                                    # Fallback 1: click vía JS
                                    try:
                                        driver.execute_script("arguments[0].click();", el_shipping)
                                    except Exception:
                                        # Fallback 2: mover con Actions y click
                                        try:
                                            ActionChains(driver).move_to_element(el_shipping).click().perform()
                                        except Exception:
                                            # Fallback 3: focus y enviar ENTER
                                            try:
                                                driver.execute_script("arguments[0].focus();", el_shipping)
                                                el_shipping.send_keys(Keys.ENTER)
                                            except Exception as final_err:
                                                print("No se pudo hacer click en Shipping:", final_err)

                            except TimeoutException:
                                print("No se encontró el tab Shipping en la página.")

                            # Localiza el campo de entrada por su ID
                            input_field = WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                            )

                            # Asegurar visibilidad antes de click
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                            time.sleep(0.3)

                            current_value = input_field.get_attribute('value')

                            if current_value != tracking:
                                try:
                                    input_field.clear()

                                    input_field.send_keys(tracking)
                                except Exception:
                                    # fallback por si el click normal falla
                                    driver.execute_script(
                                        "arguments[0].value = arguments[1];"
                                        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                                        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                                        input_field, tracking
                                    )

                                # Guardar progreso
                                button_save_clon = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located((By.ID, 'save-order-btn'))).click()

                                time.sleep(5)

                                # Entrar a PO y accionar
                                invoice_url = f"https://use1.omsapp.com/patt-op.php?scode=invoice&oID={so}"
                                driver.get(invoice_url)

                                button_Custom = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located(
                                        (By.XPATH, "//span[contains(text(), 'Custom fields')]")))
                                button_Custom.click()

                                xpath_shipping = "//ul[@id='custom-fields-tabs']//a[normalize-space(.)='Shipping']"

                                try:
                                    el_shipping = WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located((By.XPATH, xpath_shipping))
                                    )

                                    # Traer a vista
                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});",
                                                          el_shipping)
                                    time.sleep(0.2)

                                    # Intento normal esperando que sea clickeable
                                    try:
                                        WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((By.XPATH, xpath_shipping)))
                                        el_shipping.click()
                                    except (ElementNotInteractableException, ElementClickInterceptedException,
                                            TimeoutException):
                                        # Fallback 1: click vía JS
                                        try:
                                            driver.execute_script("arguments[0].click();", el_shipping)
                                        except Exception:
                                            # Fallback 2: mover con Actions y click
                                            try:
                                                ActionChains(driver).move_to_element(el_shipping).click().perform()
                                            except Exception:
                                                # Fallback 3: focus y enviar ENTER
                                                try:
                                                    driver.execute_script("arguments[0].focus();", el_shipping)
                                                    el_shipping.send_keys(Keys.ENTER)
                                                except Exception as final_err:
                                                    print("No se pudo hacer click en Shipping:", final_err)

                                except TimeoutException:
                                    print("No se encontró el tab Shipping en la página.")

                                # Localiza el campo de entrada por su ID
                                input_field = WebDriverWait(driver, 30).until(
                                    EC.presence_of_element_located((By.ID, "PCF_AFTERSHI"))
                                )

                                # Asegurar visibilidad antes de click
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
                                time.sleep(0.3)

                                current_value = input_field.get_attribute('value')

                                if current_value != tracking:
                                    print(f"Orden {so} procesada con ERROR.")
                                    log_text.insert(tk.END, f"Orden {so} procesada con ERROR en track.\n")

                                    if tracking:
                                        historial.update_cell(idx, 8, tracking)
                                    continue
                                else:
                                    print(f"Orden {so} procesada correctamente.")

                                    if tracking:
                                        historial.update_cell(idx, 8, tracking)

                            else:
                                print(f"Orden {so} procesada correctamente.")
                                if tracking:
                                    historial.update_cell(idx, 8, tracking)

                        except Exception as e:
                            print(f"Error al revisar track en orden {so}", str(e))
                            log_text.insert(tk.END, f"Error al revisar track en orden {so}\n")

                            registrar_error(
                                error_message=f"Error: {str(e)}",
                                agente=agente_inconv_var.get(),
                                order_id=ref,
                                exception=e
                            )

                            if tracking:
                                historial.update_cell(idx, 8, tracking)
                            continue


                    except Exception as e:
                        print(f"Error al realizar orden {so}", str(e))
                        log_text.insert(tk.END, f"Error al realizar orden {so}\n")

                        registrar_error(
                            error_message=f"Error: {str(e)}",
                            agente=agente_inconv_var.get(),
                            order_id=ref,
                            exception=e
                        )
                        if tracking:
                            historial.update_cell(idx, 8, tracking)
                        continue

    except Exception as e:
        print("Error al realizar acción", str(e))
        log_text.insert(tk.END, f"Error al realizar acción\n")

        registrar_error(
            error_message=f"Error: {str(e)}",
            agente=agente_inconv_var.get(),
            order_id="Error general",
            exception=e
        )

    finally:
        registrar_accion(
            total_orders=ordenes_procesadas,  
            agente=agente_inconv_var.get(),
        )

        # Mostrar mensaje de éxito
        log_text.insert(tk.END, f"Finalizado. Procesadas: {ordenes_procesadas} órdenes\n")
        driver.quit()

# Declarando las variables globales
claim_options_var = None
checkboxes = []
checkbox_vars = []
quantity_vars = []
quantity_spinboxes = []
usuario_redshift = None
contrasena_redshift = None
invoice_id_entry = None
claim_options_var = None
usuario_db_entry = None
password_db_entry = None
usuario_redshift_var,contrasena_redshift_var = None, None
checkboxes = []
checkbox_vars = []
quantity_spinboxes = []
quantity_vars = []
spinbox = None
PO_completa_var = None
cancelacion_solicitada_var = None
motivos_return_var = None
motivos_inconv_var = None
cancelacion_exitosa_var = None
agente_inconv_var = None
SO_inconv_var = None
cys_options_var = None
status_label = None
custom_text_entry = None
custom_text_so = None
robot_options_var = None
log_text = None
date_entry = None
usuario_var = None
cancelacion_solicitada_parther2_var = None
spinbox = None
PO_completa_var = None
cancelacion_solicitada_var = None
motivos_return_var = None
motivos_inconv_var = None
cancelacion_exitosa_var = None
agente_inconv_var = None
SO_inconv_var = None
cys_options_var = None
status_label = None
custom_text_entry = None
custom_text_so = None
robot_options_var = None
log_text = None
valores_coma_separada_columna1 = None
valores_coma_separada_columna2 = None
valores_coma_separada_columna3 = None
valores_coma_separada_columna4 = None
valores_coma_separada_columna5 = None
valores_coma_separada_columna6 = None
valores_no_vacios_5 = None
valores_iterables = None
username_entry = None
password_entry = None
frame_left = None
frame_middle = None
frame_right = None
frame_right_2 = None
usuario = None
contrasena = None
usuario_db = None
contrasena_db = None
usuario_company = None
contrasena_company = None
usuario_company = None
contrasena_company = None
username_company_entry = None
password_company_entry = None
username_company_entry = None
contrasena_redshift_entry = None
lote = None
fecha_inicio_var, fecha_fin_var, courier_var, manifesto_var = None, None, None, None

def update_options():
    global spinbox, usuario_redshift_var,contrasena_redshift_var,checkboxes,usuario_db_entry,password_db_entry, usuario_db, contrasena_db,  password_entry, frame_left, frame_right_2, frame_right, frame_middle, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, checkbox_vars, quantity_spinboxes, quantity_vars
    # Destruye los marcos existentes y todos sus widgets
    for widget in frame_left.winfo_children():
        widget.destroy()
    for widget in frame_right.winfo_children():
        widget.destroy()
    for widget in frame_middle.winfo_children():
        widget.destroy()
    for widget in frame_right_2.winfo_children():
        widget.destroy()

    action = action_var.get()
    function_name = f"create_{action}_options"
    if function_name in globals():
        globals()[function_name]()
    else:
        print(f"Function {function_name} not found.")
        create_login_exitoso_options()

def accionable(include_execute_button=True):
    global action_var, SO_inconv_var,usuario_redshift, contrasena_redshift,lote,usuario_redshift_var,contrasena_redshift_var, usuario_db, contrasena_db,  password_entry,usuario_db_entry,password_db_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    select_action_label = tk.Label(frame_left, text="Seleccionar acción:", font=("Arial", 12))
    select_action_label.pack(pady=5)

    cupones_checkbox = tk.Radiobutton(frame_left, text="Dejar Notas oms + Drive para Courier6", variable=action_var, value="asignar_Courier6",
                                   command=update_options)
    cupones_checkbox.pack(anchor="w")

    extraer_pdf_checkboxs = tk.Radiobutton(frame_left, text="Extraer nombre PDF en descargas", variable=action_var, value="extraer_pdf",
                                      command=update_options)
    extraer_pdf_checkboxs.pack(anchor="w")

    extraer_pdf_checkbox = tk.Radiobutton(frame_left, text="Extraer nombre PDF y comparar en sheet", variable=action_var,
                                          value="extraer_nombre_pdf",
                                          command=update_options)
    extraer_pdf_checkbox.pack(anchor="w")

    facturas_Courier7_checkbox = tk.Radiobutton(frame_left, text="Facturas Courier7", variable=action_var,
                                          value="facturas_Courier7",
                                          command=update_options)
    facturas_Courier7_checkbox.pack(anchor="w")

    Courier8_checkbox = tk.Radiobutton(frame_left, text="Asignar Mail Courier8", variable=action_var,
                                              value="Courier8",
                                              command=update_options)
    Courier8_checkbox.pack(anchor="w")

    track_Courier8_checkbox = tk.Radiobutton(frame_left, text="Asignar Tracks para MA", variable=action_var,
                                           value="asignar_track_MA",
                                           command=update_options)
    track_Courier8_checkbox.pack(anchor="w")

    facturas_Courier4_checkbox = tk.Radiobutton(frame_left, text="Facturas Courier4", variable=action_var,
                                          value="facturas_Courier4",
                                          command=update_options)
    facturas_Courier4_checkbox.pack(anchor="w")

    cuil_checkbox = tk.Radiobutton(frame_left, text="Corregir CUIL/CUIT", variable=action_var,
                                          value="cuil",
                                          command=update_options)
    cuil_checkbox.pack(anchor="w")

    facturas_oms_Courier4_checkbox = tk.Radiobutton(frame_left, text="Facturas oms Courier4", variable=action_var,
                                          value="facturas_oms_Courier4",
                                          command=update_options)
    facturas_oms_Courier4_checkbox.pack(anchor="w")

    company_refurbish_checkbox = tk.Radiobutton(frame_left, text="company Refurbish (seguimiento y template)", variable=action_var,
                                            value="company_refurbish",
                                            command=update_options)
    company_refurbish_checkbox.pack(anchor="w")

    taxes_option = tk.Radiobutton(frame_left, text="Taxes", variable=action_var,
                                            value="taxes",
                                            command=update_options)
    taxes_option.pack(anchor="w")

    pretaxes_option = tk.Radiobutton(frame_left, text="Pre Taxes", variable=action_var,
                                             value="pretaxes",
                                             command=update_options)
    pretaxes_option.pack(anchor="w")

    descripciones_cr_option = tk.Radiobutton(frame_left, text="Descripciones CR (PDF por orden)", variable=action_var,
                                             value="descripciones_cr",
                                             command=update_options)
    descripciones_cr_option.pack(anchor="w")

    login_checkbox = tk.Radiobutton(frame_left, text="Login", variable=action_var, value="login",
                                    command=update_options)
    login_checkbox.pack(anchor="w")

    if include_execute_button:
        # Crear botones de opción para agente
        agente_inconv_label = tk.Label(frame_left, text="Agente:", font=("Arial", 12))
        agente_inconv_label.pack(pady=5)

        agente_inconv_options = [
            "User 1",
            "User 2",
            "ADMIN"
        ]

        agente_inconv_var = tk.StringVar(root)
        agente_inconv_var.set(agente_inconv_options[0])  # Valor por defecto

        agente_inconv_menu = tk.OptionMenu(frame_left, agente_inconv_var, *agente_inconv_options)
        agente_inconv_menu.pack(pady=5)

        # Widget de texto para mostrar el registro
        log_text = tk.Text(frame_middle, wrap="word", height=30, width=50)
        log_text.pack(side="top", fill="both", expand=True)
        log_text.insert(tk.END, "Registro de acciones:\n")

        # Botón para ejecutar el código
        execute_button = tk.Button(frame_middle, text="Ejecutar", command=lambda: execute_code(status_label),
                                   font=("Arial", 9))
        execute_button.pack(pady=10, padx=100)

        # Etiqueta para mostrar el estado de la acción
        status_label = tk.Label(frame_middle, text="", font=("Arial", 12))
        status_label.pack(side="bottom", anchor="se", padx=20, pady=20)

def create_facturas_Courier4_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_descripciones_cr_options():
    global action_var, frame_left, frame_right_2, frame_right, frame_middle, log_text, checkboxes, checkbox_vars, spinbox, status_label, select_option_var, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_taxes_options():
    global action_var,fecha_inicio_var,usuario_redshift, contrasena_redshift, fecha_fin_var, courier_var, SO_inconv_var, lote, usuario_db, contrasena_db, password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars
    root.geometry("800x680")

    accionable()


    # Importar DateEntry de tkcalendar
    try:
        from tkcalendar import DateEntry
    except ImportError:
        import sys
        import suomsrocess
        suomsrocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
        from tkcalendar import DateEntry

    fecha_inicio_label = tk.Label(frame_left, text="Fecha inicio:", font=("Arial", 11))
    fecha_inicio_label.pack(pady=2)
    fecha_inicio_var = tk.StringVar(root)
    fecha_inicio_entry = DateEntry(frame_left, textvariable=fecha_inicio_var, width=15, font=("Arial", 11), date_pattern='yyyy-mm-dd', locale='es_ES')
    fecha_inicio_entry.pack(pady=2)

    fecha_fin_label = tk.Label(frame_left, text="Fecha fin:", font=("Arial", 11))
    fecha_fin_label.pack(pady=2)
    fecha_fin_var = tk.StringVar(root)
    fecha_fin_entry = DateEntry(frame_left, textvariable=fecha_fin_var, width=15, font=("Arial", 11), date_pattern='yyyy-mm-dd', locale='es_ES')
    fecha_fin_entry.pack(pady=2)

    courier_label = tk.Label(frame_left, text="Courier:", font=("Arial", 11))
    courier_label.pack(pady=2)
    courier_var = tk.StringVar(root)
    courier_options = ["", "Courier1", "Courier2", "Courier3", "Courier4", "Courier5"]
    courier_var.set(courier_options[0])
    courier_menu = tk.OptionMenu(frame_left, courier_var, *courier_options)
    courier_menu.pack(pady=2)

def create_pretaxes_options():
    global action_var, courier_var, manifesto_var, usuario_redshift, contrasena_redshift, \
        SO_inconv_var, lote, usuario_db, contrasena_db, password_entry, username_entry, \
        username_company_entry, password_company_entry, usuario_company, \
        contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, \
        invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, \
        PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, \
        cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, \
        status_label, custom_text_entry, custom_text_so, select_option_var, \
        quantity_spinboxes, quantity_vars

    root.geometry("800x680")
    accionable()

    courier_label = tk.Label(frame_left, text="Courier:", font=("Arial", 11))
    courier_label.pack(pady=2)
    courier_var = tk.StringVar(root)
    courier_options = ["", "Courier1", "Courier2", "Courier3", "Courier4", "Courier5"]
    courier_var.set(courier_options[0])
    courier_menu = tk.OptionMenu(frame_left, courier_var, *courier_options)
    courier_menu.pack(pady=2)

    manifesto_label = tk.Label(frame_left, text="Nro. Manifiesto:", font=("Arial", 11))
    manifesto_label.pack(pady=2)
    manifesto_var = tk.StringVar(root)
    manifesto_entry = tk.Entry(frame_left, textvariable=manifesto_var, width=20, font=("Arial", 11))
    manifesto_entry.pack(pady=2)

def create_company_refurbish_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_facturas_oms_Courier4_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_cuil_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_asignar_track_MA_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_asignar_Courier6_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

    # Crear botones de opción para lote
    lote_inconv_label = tk.Label(frame_left, text="Lote:", font=("Arial", 12))
    lote_inconv_label.pack(pady=5)

    lote_inconv_options = [
        "1",
        "2"
    ]

    lote = tk.StringVar(root)
    lote.set(lote_inconv_options[0])  # Valor por defecto

    lote_inconv_menu = tk.OptionMenu(frame_left, lote, *lote_inconv_options)
    lote_inconv_menu.pack(pady=5)

def create_Courier8_options():
    global action_var, SO_inconv_var,lote,usuario_db_entry, usuario_db, contrasena_db, password_db_entry, password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company_entry, password_company_entry, usuario_company, contrasena_company, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, method_refund, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, password_entry, usuario_var, invoice_id_list, invoice_id_entry, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_extraer_pdf_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_facturas_Courier7_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def create_extraer_nombre_pdf_options():
    global action_var, SO_inconv_var,lote, usuario_db, contrasena_db,  password_entry, username_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, frame_left, frame_right_2, frame_right, frame_middle, invoice_id_list, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var, spinbox, checkboxes, checkbox_vars, quantity_spinboxes, quantity_vars

    root.geometry("800x680")

    accionable()

def show_password_temporarily():
    if password_entry:
        password_entry.config(show='')  # Mostrar contraseña
        root.after(1000, lambda: password_entry.config(show='*'))
    if password_company_entry:
        password_company_entry.config(show='')  # Mostrar contraseña
        root.after(1000, lambda: password_company_entry.config(show='*'))
    if 'password_company_entry' in globals() and password_company_entry:
        password_company_entry.config(show='')  # Mostrar contraseña
        root.after(1000, lambda: password_company_entry.config(show='*'))
    if password_db_entry:
        password_db_entry.config(show='')  # Mostrar contraseña
        root.after(1000, lambda: password_db_entry.config(show='*'))
    if 'contrasena_redshift_entry' in globals() and contrasena_redshift_entry:
        contrasena_redshift_entry.config(show='') 
        root.after(1000, lambda: contrasena_redshift_entry.config(show='*'))

def create_login_exitoso_options():
    global action_var, username_entry, usuario_db, contrasena_db,  username_company_entry, password_company_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, method_refund, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, password_entry, usuario_var, invoice_id_list, invoice_id_entry, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var
    root.geometry("800x680")

    accionable(include_execute_button=False)

    status_label = tk.Label(frame_middle, text="Login exitoso", font=("Arial", 12))
    status_label.pack(pady=5)

    # Crear el label con el emoji
    true_label = tk.Label(frame_middle, text="✅", font=("Arial", 24))
    true_label.pack(pady=5)

def create_login_options():
    global action_var, username_entry, usuario_redshift, contrasena_redshift, usuario_db, usuario_redshift_var,contrasena_redshift_var,contrasena_db, usuario_db_entry,password_db_entry, log_text, username_company_entry, password_company_entry, username_company_entry, password_company_entry, usuario_company, contrasena_company, usuario_company, contrasena_company, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, username_entry, log_text, usuario, contrasena, frame_left, frame_right_2, frame_right, frame_middle, password_entry, method_refund, usuario_var, invoice_id_list, invoice_id_entry, claim_options_var, log_text, checkboxes, checkbox_vars, spinbox, PO_completa_var, cancelacion_solicitada_var, motivos_return_var, motivos_inconv_var, cancelacion_exitosa_var, agente_inconv_var, SO_inconv_var, cys_options_var, status_label, custom_text_entry, custom_text_so, select_option_var
    root.geometry("800x880")

    accionable(include_execute_button=False)

    select_action_label = tk.Label(frame_middle, text="Cuenta de oms:", font=("Arial", 12))
    select_action_label.pack(pady=5)

    username = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
    username.pack(fill='x', pady=5)

    username_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
    username_entry.pack(pady=5, anchor='w')
    username_entry.insert(0, '@company.com')

    password = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
    password.pack(pady=5, fill='x')

    password_frame = tk.Frame(frame_middle)
    password_frame.pack(anchor='w')

    password_entry = tk.Entry(password_frame, width=20, font=("Arial", 12), show="*")
    password_entry.pack(side="left", padx=5, anchor='w')

    toggle_button = tk.Button(password_frame, text="👁️", command=show_password_temporarily, borderwidth=0,
                              highlightthickness=0, font=("Arial", 16))
    toggle_button.pack(side="left", padx=5, anchor='w')

    company_action_label = tk.Label(frame_middle, text="Cuenta de Mail Courier8 (opcional):", font=("Arial", 12))
    company_action_label.pack(pady=5)

    username_company = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
    username_company.pack(fill='x', pady=5)

    username_company_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
    username_company_entry.pack(pady=5, anchor='w')

    password_company = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
    password_company.pack(pady=5, fill='x')

    password_company_frame = tk.Frame(frame_middle)
    password_company_frame.pack(anchor='w')

    password_company_entry = tk.Entry(password_company_frame, width=20, font=("Arial", 12), show="*")
    password_company_entry.pack(side="left", padx=5, anchor='w')

    toggle_button_company = tk.Button(password_company_frame, text="👁️", command=show_password_temporarily,
                                        borderwidth=0, highlightthickness=0, font=("Arial", 16))
    toggle_button_company.pack(side="left", padx=5, anchor='w')

    # Cuenta company (opcional)
    company_action_label = tk.Label(frame_middle, text="Cuenta de company (opcional):", font=("Arial", 12))
    company_action_label.pack(pady=5)

    username_company = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
    username_company.pack(fill='x', pady=5)

    username_company_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
    username_company_entry.pack(pady=5, anchor='w')

    password_company = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
    password_company.pack(pady=5, fill='x')

    password_company_frame = tk.Frame(frame_middle)
    password_company_frame.pack(anchor='w')

    password_company_entry = tk.Entry(password_company_frame, width=20, font=("Arial", 12), show="*")
    password_company_entry.pack(side="left", padx=5, anchor='w')

    toggle_button_company = tk.Button(password_company_frame, text="👁️", command=show_password_temporarily,
                                     borderwidth=0, highlightthickness=0, font=("Arial", 16))
    toggle_button_company.pack(side="left", padx=5, anchor='w')

    usuario_redshift_label = tk.Label(frame_right, text="Usuario Redshift (opcional):", font=("Arial", 12), anchor='w')
    usuario_redshift_label.pack(fill='x', pady=5)

    usuario_redshift_var = tk.Entry(frame_right, width=20, font=("Arial", 12))
    usuario_redshift_var.pack(pady=5, anchor='w')

    contrasena_redshift_var = tk.Label(frame_right, text="Contraseña Redshift:", font=("Arial", 12), anchor='w')
    contrasena_redshift_var.pack(fill='x', pady=5)

    contrasena_redshift_entry = tk.Entry(frame_right, width=20, font=("Arial", 12), show="*")
    contrasena_redshift_entry.pack(pady=5, anchor='w')

    toggle_button_redshift = tk.Button(frame_right, text="👁️", command=show_password_temporarily
                                       , borderwidth=0, highlightthickness=0, font=("Arial", 16))
    toggle_button_redshift.pack(side="left", padx=5, anchor='w')


    execute_button = tk.Button(frame_middle, text="Siguiente", command=lambda: login(status_label), font=("Arial", 9))
    execute_button.pack(pady=10, anchor='w')

    status_label = tk.Label(frame_middle, text="", font=("Arial", 12))
    status_label.pack(side="bottom", anchor="se", padx=20, pady=20)

root = tk.Tk()
root.title("Bots Logistica")

# Calcular la posición x e y para centrar la ventana
window_width = 550
window_height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = int((screen_width - window_width) / 2)
position_y = int((screen_height - window_height) / 2)
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# Creando los marcos para cada columna
frame_left = tk.Frame(root)
frame_left.pack(side="left", padx=20, pady=30)

frame_middle = tk.Frame(root)
frame_middle.pack(side="left", padx=20, pady=30)

frame_right = tk.Frame(root)
frame_right.pack(side="left", padx=20, pady=30)

frame_right_2 = tk.Frame(root)
frame_right_2.pack(side="left", padx=20, pady=30)

# Declarando las variables globales
action_var = tk.StringVar(root)

select_action_label = tk.Label(frame_middle, text="Cuenta de oms:", font=("Arial", 12))
select_action_label.pack(pady=5)

username = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
username.pack(fill='x', pady=5)

username_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
username_entry.pack(pady=5, anchor='w')
username_entry.insert(0, '@company.com')

password = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
password.pack(pady=5, fill='x')

password_frame = tk.Frame(frame_middle)
password_frame.pack(anchor='w')

password_entry = tk.Entry(password_frame, width=20, font=("Arial", 12), show="*")
password_entry.pack(side="left", padx=5, anchor='w')

toggle_button = tk.Button(password_frame, text="👁️", command=show_password_temporarily, borderwidth=0,
                          highlightthickness=0, font=("Arial", 16))
toggle_button.pack(side="left", padx=5, anchor='w')

company_action_label = tk.Label(frame_middle, text="Cuenta de Mail Courier8 (opcional):", font=("Arial", 12))
company_action_label.pack(pady=5)

username_company = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
username_company.pack(fill='x', pady=5)

username_company_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
username_company_entry.pack(pady=5, anchor='w')

password_company = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
password_company.pack(pady=5, fill='x')

password_company_frame = tk.Frame(frame_middle)
password_company_frame.pack(anchor='w')

password_company_entry = tk.Entry(password_company_frame, width=20, font=("Arial", 12), show="*")
password_company_entry.pack(side="left", padx=5, anchor='w')

toggle_button_company = tk.Button(password_company_frame, text="👁️", command=show_password_temporarily,
                                    borderwidth=0, highlightthickness=0, font=("Arial", 16))
toggle_button_company.pack(side="left", padx=5, anchor='w')

# Cuenta company (opcional)
company_action_label = tk.Label(frame_middle, text="Cuenta de company (opcional):", font=("Arial", 12))
company_action_label.pack(pady=5)

username_company = tk.Label(frame_middle, text="Usuario:", font=("Arial", 12), anchor='w')
username_company.pack(fill='x', pady=5)

username_company_entry = tk.Entry(frame_middle, width=20, font=("Arial", 12))
username_company_entry.pack(pady=5, anchor='w')

password_company = tk.Label(frame_middle, text="Contraseña:", font=("Arial", 12), anchor='w')
password_company.pack(pady=5, fill='x')

password_company_frame = tk.Frame(frame_middle)
password_company_frame.pack(anchor='w')

password_company_entry = tk.Entry(password_company_frame, width=20, font=("Arial", 12), show="*")
password_company_entry.pack(side="left", padx=5, anchor='w')

toggle_button_company = tk.Button(password_company_frame, text="👁️", command=show_password_temporarily,
                                 borderwidth=0, highlightthickness=0, font=("Arial", 16))
toggle_button_company.pack(side="left", padx=5, anchor='w')

usuario_redshift_label = tk.Label(frame_right, text="Usuario Redshift (opcional):", font=("Arial", 12), anchor='w')
usuario_redshift_label.pack(fill='x', pady=5)

usuario_redshift_var = tk.Entry(frame_right, width=20, font=("Arial", 12))
usuario_redshift_var.pack(pady=5, anchor='w')

contrasena_redshift_var = tk.Label(frame_right, text="Contraseña Redshift:", font=("Arial", 12), anchor='w')
contrasena_redshift_var.pack(fill='x', pady=5)

contrasena_redshift_entry = tk.Entry(frame_right, width=20, font=("Arial", 12), show="*")
contrasena_redshift_entry.pack(pady=5, anchor='w')

toggle_button_redshift = tk.Button(frame_right, text="👁️", command=show_password_temporarily
                                    , borderwidth=0, highlightthickness=0, font=("Arial", 16))
toggle_button_redshift.pack(side="left", padx=5, anchor='w')


execute_button = tk.Button(frame_middle, text="Siguiente", command=lambda: login(status_label), font=("Arial", 9))
execute_button.pack(pady=10, anchor='w')

status_label = tk.Label(frame_middle, text="", font=("Arial", 12))
status_label.pack(side="bottom", anchor="se", padx=20, pady=20)

root.mainloop()


