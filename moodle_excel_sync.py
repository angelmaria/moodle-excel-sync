#!/usr/bin/env python3
"""
Script simple para crear/editar usuarios en Moodle desde Excel
"""

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime

EXCEL_FILE = '/Users/cash/Desktop/TomasMelendo/Registro_curso_Amor_sexualidad.xlsx'
LOG_FILE = '/Users/cash/Desktop/TomasMelendo/log_creacion_usuarios.txt'

# ===== CONFIGURACIÓN DE FILAS A PROCESAR =====
# Define qué filas del Excel procesará el script
# Opción 1: Rango continuo
FILAS_A_PROCESAR = list(range(181, 285))  # Filas 181 a 284

# Opción 2: Filas específicas (descomenta para usar)
# FILAS_A_PROCESAR = [177, 178, 179, 180]

# Opción 3: Reintentos de filas específicas (descomenta para usar)
# FILAS_A_PROCESAR = [254, 275]

# Opción 4: Una sola fila
# FILAS_A_PROCESAR = [180]
# ============================================

def log_msg(mensaje):
    """Imprime mensaje con timestamp tanto en consola como en log"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linea = f"[{ts}] {mensaje}"
    print(mensaje)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(linea + '\n')

def leer_registros_excel(filas):
    """Lee los datos de los registros especificados del Excel"""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    
    registros = []
    for fila in filas:
        apellidos = ws.cell(fila, 1).value  # Columna 1: Apellidos
        nombre = ws.cell(fila, 2).value     # Columna 2: Nombre
        email = ws.cell(fila, 3).value      # Columna 3: Email
        usuario = ws.cell(fila, 6).value    # Columna 6: Usuario
        contrasena = ws.cell(fila, 7).value # Columna 7: Contraseña
        
        if nombre and apellidos and email and usuario:
            registros.append({
                'fila': fila,
                'nombre': nombre.strip(),
                'apellidos': apellidos.strip(),
                'email': email.strip(),
                'usuario': usuario.strip().lower(),
                'contrasena': contrasena.strip() if contrasena else None
            })
    
    return registros

def login_moodle(driver):
    """Inicia sesión en Moodle"""
    log_msg("\n✓ Iniciando sesión en Moodle...")
    
    driver.get("https://campus.edufamilia.com/admin/search.php")
    time.sleep(2)
    
    wait = WebDriverWait(driver, 15)
    
    # Usuario
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "username")))
    campo_usuario.send_keys("admin")
    
    # Contraseña
    campo_password = driver.find_element(By.ID, "password")
    campo_password.send_keys("ax%$85-.BXD")
    
    # Click en Log in
    login_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")
    login_btn.click()
    
    time.sleep(4)
    log_msg("✓ Sesión iniciada")

def procesar_usuario(driver, registro, es_primero=False):
    """Verifica si el usuario existe por email y lo crea o edita según corresponda"""
    fila = registro['fila']
    nombre = registro['nombre']
    apellidos = registro['apellidos']
    email = registro['email']
    usuario = registro['usuario']
    contrasena = registro['contrasena']
    
    log_msg(f"\n[Fila {fila}] Procesando: {nombre} {apellidos} ({email})")
    
    wait = WebDriverWait(driver, 15)
    
    try:
        # 1. Ir a "Examinar lista de usuarios"
        driver.get("https://campus.edufamilia.com/admin/user.php")
        time.sleep(2)
        
        # 2. Si NO es el primero, eliminar filtros anteriores
        if not es_primero:
            try:
                eliminar_filtros = wait.until(EC.element_to_be_clickable((By.ID, "id_removeall")))
                eliminar_filtros.click()
                time.sleep(1)
            except:
                pass  # Si no hay filtros, continuar
        
        # 3. Hacer clic en "Mostrar más..."
        mostrar_mas = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.moreless-toggler")))
        mostrar_mas.click()
        time.sleep(1)
        
        # 4. Buscar por email
        campo_email = wait.until(EC.presence_of_element_located((By.ID, "id_email")))
        campo_email.clear()
        campo_email.send_keys(email)
        campo_email.send_keys(Keys.RETURN)
        time.sleep(3)
        
        # 5. Verificar si encuentra usuarios
        try:
            # Verificar si dice "No se encuentran usuarios"
            no_encontrado = driver.find_element(By.XPATH, "//*[contains(text(), 'No se encuentran usuarios')]")
            log_msg(f"  → Usuario NO existe. Creando...")
            
            # Hacer clic en "Crear un nuevo usuario"
            boton_crear = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Crear un nuevo usuario')]")))
            boton_crear.click()
            time.sleep(2)
            
            # Llamar a la función de crear usuario (ya estamos en el formulario)
            return crear_usuario_en_formulario(driver, registro)
            
        except:
            # Si no encuentra el mensaje "No se encuentran usuarios", significa que SÍ existe
            log_msg(f"  → Usuario YA existe. Editando...")
            
            # Hacer clic en el icono de configuración (engranaje) para editar
            edit_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='user/editadvanced.php'] i.fa-cog")))
            edit_icon.click()
            time.sleep(2)
            
            # Modificar Nombre
            campo_nombre = wait.until(EC.presence_of_element_located((By.ID, "id_firstname")))
            campo_nombre.clear()
            campo_nombre.send_keys(nombre)
            
            # Modificar Apellidos
            campo_apellido = driver.find_element(By.ID, "id_lastname")
            campo_apellido.clear()
            campo_apellido.send_keys(apellidos)
            
            # Guardar cambios
            boton_guardar = driver.find_element(By.ID, "id_submitbutton")
            driver.execute_script("arguments[0].scrollIntoView(true);", boton_guardar)
            time.sleep(1)
            boton_guardar.click()
            time.sleep(3)
            
            log_msg(f"  ✓✓ Usuario editado exitosamente")
            return True
            
    except Exception as e:
        log_msg(f"  ✗ Error: {str(e)[:100]}")
        return False

def crear_usuario_en_formulario(driver, registro):
    """Crea un usuario cuando ya estamos en el formulario de creación"""
    try:
        wait = WebDriverWait(driver, 15)
        
        # 1. Esperar y rellenar Nombre de usuario
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "id_username")))
        campo_usuario.clear()
        campo_usuario.send_keys(registro['usuario'])
        time.sleep(0.5)
        log_msg(f"  Nombre de usuario: {registro['usuario']}")
        
        # 2. Rellenar Dirección de correo
        campo_email = driver.find_element(By.ID, "id_email")
        campo_email.clear()
        campo_email.send_keys(registro['email'])
        time.sleep(0.5)
        log_msg(f"  Email: {registro['email']}")
        
        # 3. Rellenar Nombre
        campo_nombre = driver.find_element(By.ID, "id_firstname")
        campo_nombre.clear()
        campo_nombre.send_keys(registro['nombre'])
        time.sleep(0.5)
        log_msg(f"  Nombre: {registro['nombre']}")
        
        # 4. Rellenar Apellidos
        campo_apellido = driver.find_element(By.ID, "id_lastname")
        campo_apellido.clear()
        campo_apellido.send_keys(registro['apellidos'])
        time.sleep(0.5)
        log_msg(f"  Apellidos: {registro['apellidos']}")
        
        # 5. Hacer clic en "Haz click para insertar texto" para la contraseña
        log_msg(f"  ⏳ Buscando enlace de contraseña...")
        try:
            enlace_contrasena = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-passwordunmask='edit']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", enlace_contrasena)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", enlace_contrasena)
            log_msg(f"  ✓ Click en enlace de contraseña")
            time.sleep(2)
        except Exception as e:
            log_msg(f"  ⚠ Error con enlace de contraseña: {str(e)[:50]}")
            pass
        
        # 6. Rellenar Nueva contraseña
        try:
            campo_password = wait.until(EC.presence_of_element_located((By.ID, "id_newpassword")))
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_password)
            time.sleep(1)
            campo_password.clear()
            campo_password.send_keys(registro['contrasena'])
            time.sleep(0.5)
            log_msg(f"  Contraseña: {registro['contrasena']}")
        except Exception as e:
            log_msg(f"  ✗ Error al rellenar contraseña: {str(e)[:50]}")
            raise
        
        # 7. Hacer clic en "Crear usuario"
        submit_btn = wait.until(EC.element_to_be_clickable((By.NAME, "submitbutton")))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        time.sleep(1)
        submit_btn.click()
        log_msg(f"  ✓ Haciendo click en 'Crear Usuario'")
        time.sleep(3)
        
        # 8. Verificar error
        try:
            error = driver.find_element(By.CLASS_NAME, "alert-danger")
            error_text = error.text
            log_msg(f"  ✗ Error: {error_text[:80]}")
            return False
        except:
            log_msg(f"  ✓✓ Usuario creado exitosamente")
            return True
            
    except Exception as e:
        log_msg(f"  ✗ Error: {str(e)[:100]}")
        return False

def crear_usuario_moodle(driver, registro):
    """Crea un usuario en Moodle"""
    fila = registro['fila']
    nombre = registro['nombre']
    apellidos = registro['apellidos']
    email = registro['email']
    usuario = registro['usuario']
    contrasena = registro['contrasena']
    
    log_msg(f"\n[Fila {fila}] Creando usuario: {usuario}")
    log_msg(f"  Nombre: {nombre} | Apellidos: {apellidos} | Email: {email}")
    
    crear_usuario_url = "https://campus.edufamilia.com/user/editadvanced.php?id=-1"
    driver.get(crear_usuario_url)
    time.sleep(2)
    
    wait = WebDriverWait(driver, 15)
    
    try:
        # Campo usuario
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "id_username")))
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)
        
        # Campo email
        campo_email = driver.find_element(By.ID, "id_email")
        campo_email.clear()
        campo_email.send_keys(email)
        
        # Campo nombre
        campo_nombre = driver.find_element(By.ID, "id_firstname")
        campo_nombre.clear()
        campo_nombre.send_keys(nombre)
        
        # Campo apellido
        campo_apellido = driver.find_element(By.ID, "id_lastname")
        campo_apellido.clear()
        campo_apellido.send_keys(apellidos)
        
        # Habilitar contraseña
        try:
            enlace_pwd = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[@data-passwordunmask='edit']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", enlace_pwd)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", enlace_pwd)
            time.sleep(2)
        except:
            pass
        
        # Llenar contraseña
        try:
            campo_pwd = wait.until(EC.presence_of_element_located((By.ID, "id_newpassword")))
            time.sleep(1)
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_pwd)
            campo_pwd.send_keys(contrasena)
        except:
            log_msg(f"  ⚠ Error en contraseña")
        
        # Click en Crear Usuario
        boton_crear = wait.until(EC.element_to_be_clickable((By.NAME, "submitbutton")))
        boton_crear.click()
        time.sleep(3)
        
        # Verificar error
        try:
            error = driver.find_element(By.CLASS_NAME, "alert-danger")
            error_text = error.text
            log_msg(f"  ✗ Error: {error_text[:80]}")
            return False
        except:
            log_msg(f"  ✓✓ Usuario creado exitosamente")
            return True
            
    except Exception as e:
        log_msg(f"  ✗ Error: {str(e)[:100]}")
        return False

def editar_usuario_moodle(driver, fila, nombre, apellidos, email):
    """Edita un usuario existente en Moodle buscando por email"""
    log_msg(f"\n[Fila {fila}] Editando usuario...")
    log_msg(f"  Nombre: {nombre} | Apellidos: {apellidos} | Email: {email}")
    
    wait = WebDriverWait(driver, 15)
    
    try:
        # 1. Ir a "Examinar lista de usuarios"
        driver.get("https://campus.edufamilia.com/admin/user.php")
        time.sleep(2)
        
        # 2. Hacer clic en "Mostrar más..."
        mostrar_mas = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.moreless-toggler")))
        mostrar_mas.click()
        time.sleep(1)
        
        # 3. Escribir email en el campo "Dirección de correo"
        campo_email = wait.until(EC.presence_of_element_located((By.ID, "id_email")))
        campo_email.clear()
        campo_email.send_keys(email)
        campo_email.send_keys(Keys.RETURN)
        time.sleep(3)
        
        # 4. Hacer clic en el icono de configuración (engranaje) para editar
        edit_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='user/editadvanced.php'] i.fa-cog")))
        edit_icon.click()
        time.sleep(2)
        
        # 5. Modificar Nombre
        campo_nombre = wait.until(EC.presence_of_element_located((By.ID, "id_firstname")))
        campo_nombre.clear()
        campo_nombre.send_keys(nombre)
        
        # 6. Modificar Apellidos
        campo_apellido = driver.find_element(By.ID, "id_lastname")
        campo_apellido.clear()
        campo_apellido.send_keys(apellidos)
        
        # 7. Guardar cambios
        boton_guardar = driver.find_element(By.ID, "id_submitbutton")
        driver.execute_script("arguments[0].scrollIntoView(true);", boton_guardar)
        time.sleep(1)
        boton_guardar.click()
        time.sleep(3)
        
        log_msg(f"  ✓✓ Usuario editado exitosamente")
        return True
        
    except Exception as e:
        log_msg(f"  ✗ Error: {str(e)[:100]}")
        return False

def main():
    """Función principal"""
    log_msg("=" * 80)
    log_msg("PROCESAR USUARIOS EN MOODLE")
    log_msg("=" * 80)
    
    # Configurar Chrome
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("user-agent=Mozilla/5.0")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    try:
        # Login
        login_moodle(driver)
        
        # Procesar registros según configuración
        registros = leer_registros_excel(FILAS_A_PROCESAR)
        
        log_msg(f"\nProcesando {len(registros)} registros...")
        for r in registros:
            log_msg(f"  - Fila {r['fila']}: {r['nombre']} {r['apellidos']}")
        
        for idx, registro in enumerate(registros):
            es_primero = (idx == 0)
            procesar_usuario(driver, registro, es_primero)
            time.sleep(1)
        
        log_msg("\n" + "=" * 80)
        log_msg("✓ Proceso completado")
        log_msg("=" * 80)
        
    except Exception as e:
        log_msg(f"\n✗ Error general: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
