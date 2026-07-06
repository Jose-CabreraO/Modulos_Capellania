# Modulos Capellania

Aplicacion multipagina en Streamlit para gestionar modulos operativos de Capellania. El primer modulo disponible es `Capellania`, con cola editable de grupos, ingestion desde WhatsApp y automatizacion de carga con Playwright.

## Requisitos

- Python 3.10+
- Google Chrome o Chromium para Playwright
- Archivo local `plantilla_flujo_completo.xlsx` en la raiz del proyecto

## Instalacion

```powershell
cd "C:\Users\Jose\Desktop\Grupos de estudio"
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Variables de entorno

Crea un archivo `.env` local o define estas variables en la terminal:

```powershell
$env:CAPELLANIA_USER="tu_usuario"
$env:CAPELLANIA_PASS="tu_contrasena"
```

El archivo `.env` y el Excel operativo no se suben a GitHub.

## Ejecutar la aplicacion

```powershell
python -m streamlit run app.py
```

Luego abre:

```text
http://localhost:8501
```

## Estructura

```text
app.py
core/
  bot_carga.py
  bot_referencias.py
  excel_store.py
  parser_whatsapp.py
modulos/
  capellania.py
requirements.txt
```

## Flujo principal

- La hoja `1_Crear_Grupos` mantiene el estado de carga con `Estado_Carga`.
- El bot procesa solo filas `PENDIENTE`.
- Las filas exitosas pasan a `PROCESADO` y se pintan en verde.
- Las filas con error pasan a `ERROR` y se pintan en rojo.
- La grilla de Streamlit permite corregir datos y volver a dejar filas en `PENDIENTE`.
