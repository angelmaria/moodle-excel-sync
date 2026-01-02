# Moodle Excel Sync

Herramienta automatizada para crear y actualizar usuarios en Moodle desde un archivo Excel, con detección inteligente de usuarios existentes.

## 🎯 Características

- ✅ **Lectura desde Excel**: Lee datos de usuarios directamente desde archivos `.xlsx`
- ✅ **Detección automática**: Verifica si los usuarios ya existen en Moodle por email
- ✅ **Creación masiva**: Crea nuevos usuarios con nombres y apellidos correctos
- ✅ **Actualización inteligente**: Edita usuarios existentes sin duplicarlos
- ✅ **Manejo de caracteres especiales**: Soporta acentos y caracteres Unicode
- ✅ **Logging en tiempo real**: Registro detallado de todas las operaciones
- ✅ **Automatización completa**: No requiere interacción manual

## 📋 Requisitos

- Python 3.8+
- Chrome/Chromium instalado
- Acceso administrativo a Moodle
- Archivo Excel con estructura específica

## 🚀 Instalación

```bash
# Clonar el repositorio
git clone https://github.com/tuusuario/moodle-excel-sync.git
cd moodle-excel-sync

# Crear entorno virtual
python3 -m venv .venv
source .venv/bin/activate  # En macOS/Linux
# o
.venv\Scripts\activate  # En Windows

# Instalar dependencias
pip install -r requirements.txt
```

## 📝 Estructura del Excel

El archivo Excel debe tener la siguiente estructura:

| Columna | Nombre | Contenido | Ejemplo |
|---------|--------|-----------|---------|
| 1 | Apellidos | Apellidos (pueden ser compuestos) | `Navarro Azabache` |
| 2 | Nombre | Nombres (pueden ser compuestos) | `Carlos Gabriel` |
| 3 | Email | Correo electrónico | `carlos@example.com` |
| 6 | Usuario | Nombre de usuario (único) | `carlos.gabriel` |
| 7 | Contraseña | Contraseña temporal | `Carlos+A1+-` |

**Nota**: Las columnas 4, 5 y posteriores pueden contener otros datos y serán ignoradas.

## ⚙️ Configuración

Edita las variables en `moodle_excel_sync.py`:

```python
# Ruta del archivo Excel
EXCEL_FILE = '/ruta/a/tu/archivo.xlsx'

# Ruta del archivo de log
LOG_FILE = '/ruta/a/tu/log.txt'
```

## 🎮 Uso

### Opción 1: Procesar registros específicos

```bash
python moodle_excel_sync.py
```

Por defecto procesa las filas 181-284. Para cambiar el rango, edita la línea:

```python
registros = leer_registros_excel(list(range(181, 285)))
```

### Opción 2: Procesar registros puntuales

```python
# En main(), cambiar:
registros = leer_registros_excel([177, 178, 179, 180])  # Solo estas filas
```

## 🔄 Flujo de ejecución

1. **Lectura de Excel**: Carga los datos de los registros especificados
2. **Login en Moodle**: Se autentica con credenciales de administrador
3. **Verificación por email**:
   - Si el email **existe** → Edita el usuario (actualiza nombre y apellidos)
   - Si el email **no existe** → Crea un nuevo usuario
4. **Limpieza de filtros**: Entre cada usuario, limpia los filtros anteriores
5. **Logging**: Registra todas las operaciones

## 📊 Ejemplo de ejecución

```
================================================================================
PROCESAR USUARIOS EN MOODLE
================================================================================

✓ Iniciando sesión en Moodle...
✓ Sesión iniciada

Procesando 4 registros...
  - Fila 177: Emilia Nakauchi Lago
  - Fila 178: Mayda Narvaez
  - Fila 179: Natalia Navarrete
  - Fila 180: Ana Isabel Navarro

[Fila 177] Procesando: Emilia Nakauchi Lago (emilia.nakauchi@gmail.com)
  → Usuario YA existe. Editando...
  ✓✓ Usuario editado exitosamente

[Fila 178] Procesando: Mayda Narvaez (mayda_ng@yahoo.com)
  → Usuario NO existe. Creando...
  ✓ Click en enlace de contraseña
  ✓ Haciendo click en 'Crear Usuario'
  ✓✓ Usuario creado exitosamente

================================================================================
✓ Proceso completado
================================================================================
```

## 🔐 Credenciales de Moodle

Las credenciales se configuran en el código. **IMPORTANTE**: 
- Nunca commits credenciales en el repositorio
- Usa variables de entorno o archivos `.env` en producción

## 🛠️ Solución de problemas

### Error: "Usuario NO existe" en todos
- Verifica que el campo email en Excel esté correcto
- Comprueba la conexión a Moodle
- Revisa los logs para más detalles

### Error: "invalid element state"
- El elemento está siendo procesado, reintentar generalmente funciona
- El script tiene reintentos automáticos para estos casos

### No se carga el formulario de creación
- Aumento de timeout en `WebDriverWait(driver, 15)`
- Verifica la velocidad de conexión a Moodle

## 📦 Dependencias

```
selenium==4.13.0
openpyxl==3.10.0
webdriver-manager==4.0.1
```

Ver `requirements.txt` para más detalles.

## 📄 Licencia

MIT License - libre para usar, modificar y distribuir

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:
1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📞 Soporte

Para reportar problemas o sugerencias, abre un issue en GitHub.

## 🔄 Historial de cambios

### v1.0.0 (2026-01-03)
- ✅ Versión inicial estable
- ✅ Soporte para crear y editar usuarios
- ✅ Detección automática de usuarios existentes
- ✅ Manejo completo de caracteres especiales

---

**Autor**: [Angel Martinez](mailto:angelmaria75@gmail.com)  
**Última actualización**: 3 de enero de 2026
