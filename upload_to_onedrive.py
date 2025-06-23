import os
import shutil
from datetime import datetime
import requests
import json
import glob
import time

# Archivos y configuración para el flujo de autenticación
TOKEN_FILE = "token_data.json"
SUBIDOS_FILE = "subidos.json"
CLIENT_ID = "6a9129b2-81b0-4d0d-915f-60c232db9fe2"
TENANT = "common"
SCOPES = "offline_access Files.ReadWrite.All User.Read"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
DEVICE_CODE_URL = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/devicecode"

# Guardar mensajes con hora en consola y en archivo
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")
    try:
        with open("upload_debug.log", "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"Error escribiendo log: {e}")

# Guarda el token obtenido localmente
def save_token_data(data):
    try:
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception as e:
        log_message(f"Error guardando token: {e}")

# Carga el token si ya existe guardado
def load_token_data():
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return None
    return None

# Refresca el token usando el refresh_token guardado
def refresh_access_token(refresh_token):
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": SCOPES
    }
    try:
        r = requests.post(TOKEN_URL, data=data)
        if r.status_code == 200:
            token_data = r.json()
            save_token_data(token_data)
            log_message("Token renovado correctamente.")
            return token_data["access_token"]
        else:
            log_message(f"Error renovando token: {r.text}")
            return None
    except Exception as e:
        log_message(f"Error inesperado al renovar token: {e}")
        return None

# Autenticación por código de dispositivo (si no hay refresh_token válido)
def get_access_token_device_code():
    stored = load_token_data()
    if stored and "refresh_token" in stored:
        log_message("Intentando renovar token...")
        token = refresh_access_token(stored["refresh_token"])
        if token:
            return token
        else:
            log_message("No se pudo renovar el token. Procediendo con Device Code Flow...")

    log_message("Iniciando autenticación...")
    data = {
        "client_id": CLIENT_ID,
        "scope": SCOPES
    }
    try:
        r = requests.post(DEVICE_CODE_URL, data=data)
        r.raise_for_status()
        result = r.json()
        print("\n🔑 Autoriza el acceso aquí:")
        print(f"{result['verification_uri']}")
        print(f"Código: {result['user_code']}\n")
        log_message(f"URL: {result['verification_uri']} | Código: {result['user_code']}")
        data.update({
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": result["device_code"]
        })
        # Reintenta obtener el token mientras el usuario autoriza
        while True:
            time.sleep(result["interval"])
            token_response = requests.post(TOKEN_URL, data=data)
            if token_response.status_code == 200:
                token_data = token_response.json()
                save_token_data(token_data)
                log_message("Autenticación completada con éxito.")
                return token_data["access_token"]
            elif token_response.status_code == 400:
                if token_response.json().get("error") != "authorization_pending":
                    log_message(f"Error obteniendo token: {token_response.json()}")
                    return None
            else:
                log_message(f"Error inesperado: {token_response.text}")
                return None
    except Exception as e:
        log_message(f"Error en Device Code Flow: {e}")
        return None

# Busca archivos Excel en varias rutas
def find_excel_files():
    log_message("Buscando archivos .xlsx...")
    patterns = [
        "./Data/input/*.xlsx", "./Data/*.xlsx", "./RPA/*.xlsx", "./input/*.xlsx", "./Reportes/*.xlsx"
    ]
    found = []
    for pattern in patterns:
        files = glob.glob(pattern)
        # Ignora archivos temporales de Excel
        files = [f for f in files if not os.path.basename(f).startswith('~$')]
        if files:
            log_message(f"Encontrados en {pattern}: {files}")
            found.extend(files)
    return found

# Verifica si el token funciona accediendo al Drive
def test_onedrive_connection(token):
    log_message("Verificando conexión a OneDrive...")
    headers = {'Authorization': f'Bearer {token}'}
    try:
        response = requests.get("https://graph.microsoft.com/v1.0/me/drive", headers=headers)
        response.raise_for_status()
        log_message("Conexión a OneDrive exitosa.")
        return True
    except Exception as e:
        log_message(f"Error de conexión: {e}")
        return False

# Sube el archivo a la carpeta RPA/Reportes en OneDrive
def upload_to_onedrive(file_path, token):
    log_message(f"Subiendo archivo: {file_path}")
    if not os.path.exists(file_path):
        log_message(f"Archivo no encontrado: {file_path}")
        return False

    filename = os.path.basename(file_path)
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/RPA/Reportes/{filename}:/content"
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/octet-stream'
    }
    try:
        with open(file_path, 'rb') as file:
            response = requests.put(url, headers=headers, data=file)
            if response.status_code in [200, 201]:
                log_message(f"Archivo subido correctamente: {filename}")
                return True
            else:
                log_message(f"Error en subida: {response.status_code} - {response.text}")
                return False
    except Exception as e:
        log_message(f"Error al subir archivo: {e}")
        return False

# Orquesta todo el flujo: buscar, autenticar y subir archivo
def run_full_process():
    log_message("Iniciando proceso de subida a OneDrive")
    archivos_subidos = []
    excel_files = find_excel_files()

    # Filtra solo los que tienen el formato esperado en el nombre
    reportes = [
        f for f in excel_files
        if os.path.basename(f).startswith("Reporte_") and f.endswith(".xlsx")
    ]
    if not reportes:
        log_message("No se encontró ningún archivo con formato válido.")
        return False

    # Extrae la fecha del nombre para identificar el más reciente
    def extraer_fecha(f):
        try:
            nombre = os.path.basename(f)
            fecha_txt = nombre.replace("Reporte_", "").replace(".xlsx", "")
            return datetime.strptime(fecha_txt, "%Y-%m-%d")
        except:
            return datetime.min

    archivo_mas_reciente = max(reportes, key=extraer_fecha)

    token = get_access_token_device_code()
    if not token:
        return False
    if not test_onedrive_connection(token):
        return False

    if upload_to_onedrive(archivo_mas_reciente, token):
        log_message("Proceso completado exitosamente.")
        return True
    else:
        log_message("No se pudo subir el archivo.")
        return False

# Punto de entrada
def main():
    log_message("=== EJECUCIÓN INICIADA ===")
    log_message(f"Directorio actual: {os.getcwd()}")
    try:
        run_full_process()
    except Exception as e:
        log_message(f"Error general: {e}")
    finally:
        log_message("=== EJECUCIÓN FINALIZADA ===")

# Método que puede ser invocado desde PixStudio
def ejecutar_pixstudio():
    main()
