import subprocess
import os
import pyautogui
import pyautogui as p
import time
import logging
import sys
import traceback
from datetime import datetime, UTC
from pathlib import Path
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    path = os.path.join(base_path, relative_path)
    if not os.path.exists(path):
        path = os.path.join(os.path.abspath("."), relative_path)
    return path

def run_powershell(cmd):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
        
        result = subprocess.run(
            ["powershell", "-Command", cmd],
            capture_output=True,
            text=True,
            timeout=10,
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def take_screenshot(name):
    try:
        out = Path("out")
        out.mkdir(exist_ok=True)
        ts = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
        path = out / f"{name}_{ts}.png"
        img = pyautogui.screenshot()
        img.save(path)
        logging.info(f"Captura tomada: {path}")
        return path
    except Exception as e:
        logging.error(f"Error tomando captura {name}: {str(e)}")
        return None

def fill_form(data, start_coords):
    try:
        logging.info("Iniciando llenado de formulario")
        time.sleep(3)
        take_screenshot("before")
        start_coords=622
        start_coords2=516
        try:
            pyautogui.click(start_coords, start_coords2)
            logging.info(f"Clic en coordenadas ({start_coords}, {start_coords2})")
        except Exception as e:
            logging.error(f"Error haciendo clic inicial: {str(e)}")
            return False
            
        time.sleep(1.5)
        try:
            calendario_path = get_resource_path("calendario.png")
            box = p.locateOnScreen(calendario_path, confidence=0.7)
            if box:
                p.click(p.center(box))
                logging.info("Calendario encontrado y clickeado")
            else:
                logging.warning("Calendario no encontrado")
        except Exception as e:
            logging.error(f"Error buscando calendario: {str(e)}")
            
        pyautogui.press("enter")
        pyautogui.press("tab")
        try:
            pyautogui.typewrite(data["nombres"])
            logging.info("Nombres escritos")
        except Exception as e:
            logging.error(f"Error escribiendo nombres: {str(e)}")
            
        pyautogui.press("tab")
        try:
            pyautogui.typewrite(data["matriculaperosumadaohsi"])
            logging.info("Matrícula escrita")
        except Exception as e:
            logging.error(f"Error escribiendo matrícula: {str(e)}")
            
        pyautogui.press("tab")
        try:
            opcion_path = get_resource_path("opcion.png")
            box2 = p.locateCenterOnScreen(opcion_path, confidence=0.7)
            if box2:
                p.click(box2)
                logging.info("Opción encontrada y clickeada")
            else:
                logging.warning("Opción no encontrada")
        except Exception as e:
            logging.error(f"Error buscando imagen: {str(e)}")

        take_screenshot("during")
        pyautogui.press("tab")
        pyautogui.press("enter")
        time.sleep(3)
        take_screenshot("after")
        logging.info("Formulario llenado exitosamente")
        return True
        
    except Exception as e:
        logging.error(f"Error crítico en fill_form: {str(e)}")
        take_screenshot("error_fill_form")
        return False

def main():
    try:
        logging.basicConfig(
            filename="run.log",
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            encoding="utf-8"
        )
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        console_handler.setFormatter(formatter)
        logging.getLogger().addHandler(console_handler)
        logging.info("Inicio del examen")
        try:
            screen_size = pyautogui.size()
            logging.info(f"Tamaño de pantalla detectado: {screen_size}")
        except Exception as e:
            logging.warning(f"No se pudo detectar el tamaño de pantalla: {e}")
        data = {
            "nombres": "Ian Haziel Gomez Ochoa \nSimon Wenceslao Robledo Solis \nXochilpilli Castillo Andrade",
            "matriculaperosumadaohsi": "6513716",
        }
        start_coords3 = (622, 516)
        code, out, err = run_powershell("Get-Date")
        logging.info(f"PS code: {code}")
        if out:
            logging.info(f"PS output: {out}")
        if err:
            logging.warning(f"PS error: {err}")
        time.sleep(4)
        success = fill_form(data, start_coords3)
        
        if success:
            logging.info("Terminado")
        else:
            logging.error("Errores")
            return 1

    except KeyboardInterrupt:
        logging.info("Ejecución cancelada por el usuario")
        return 130
    except Exception as e:
        logging.error(f"Error en main: {str(e)}")
        logging.error(traceback.format_exc())
        return 1
    return 0