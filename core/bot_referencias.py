import os

import pandas as pd
from playwright.sync_api import sync_playwright

from core.excel_store import EXCEL_FILE


WP_LOGIN_URL = "https://capellania-app.visualweb.systems/wp-login.php"
TARGET_URL = "https://capellania-app.visualweb.systems/wp-admin/admin.php?page=capapp2_estudios_biblicos"

USER_WP = os.getenv("CAPELLANIA_USER", "")
PASS_WP = os.getenv("CAPELLANIA_PASS", "")


def validar_credenciales():
    if not USER_WP or not PASS_WP:
        raise RuntimeError(
            "Faltan CAPELLANIA_USER y/o CAPELLANIA_PASS. "
            "Definilas como variables de entorno antes de extraer catalogos."
        )


def extraer_catalogos_completos(excel_file=EXCEL_FILE):
    validar_credenciales()
    print("Iniciando extraccion exhaustiva de catalogos desde la plataforma...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto(WP_LOGIN_URL)
        page.fill("#user_login", USER_WP)
        page.fill("#user_pass", PASS_WP)
        page.click("#wp-submit")
        page.wait_for_load_state("networkidle")

        page.goto(TARGET_URL)
        page.wait_for_load_state("networkidle")
        page.locator("table thead th:last-child").first.click()
        page.wait_for_selector("text=Grupo Estudio Biblico", timeout=5000)

        def obtener_opciones(locator_select):
            opciones = locator_select.locator("option").all_text_contents()
            return [o.strip() for o in opciones if o.strip() and "selecciona" not in o.lower() and "elija" not in o.lower()]

        capellanes = obtener_opciones(page.locator("select[name*='capellan'], select:has-text('Aguero Ever')").first)
        materiales = obtener_opciones(page.locator("select[name*='material'], select:has-text('A la manera de un Pastor')").first)
        estados = obtener_opciones(page.locator("select[name*='estado'], select:has-text('ACTIVO')").first)

        select_empresa = page.locator("select[name*='empresa'], select:has-text('FundasSA')").first
        empresas = obtener_opciones(select_empresa)
        select_sucursal = page.locator("select[name*='sucursal'], select:has-text('Casa Central')").first
        sucursales_unicas = set()

        for empresa_nombre in empresas:
            try:
                select_empresa.select_option(label=empresa_nombre)
                page.wait_for_timeout(300)
                sucursales_unicas.update(obtener_opciones(select_sucursal))
            except Exception:
                continue

        browser.close()

    sucursales = sorted(sucursales_unicas)
    max_len = max(len(empresas), len(sucursales), len(materiales), len(capellanes), len(estados), 4)
    df_listas = pd.DataFrame({
        "Empresas": empresas + [""] * (max_len - len(empresas)),
        "Sucursales": sucursales + [""] * (max_len - len(sucursales)),
        "Materiales": materiales + [""] * (max_len - len(materiales)),
        "Capellanes": capellanes + [""] * (max_len - len(capellanes)),
        "Estados": estados + [""] * (max_len - len(estados)),
        "Secciones": ["Ecommerce", "Deposito nuevo", "Administracion", "Produccion"] + [""] * (max_len - 4),
    })

    with pd.ExcelWriter(excel_file, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        df_listas.to_excel(writer, sheet_name="Listas_Referencia", index=False)

    return {
        "empresas": len(empresas),
        "sucursales": len(sucursales),
        "materiales": len(materiales),
        "capellanes": len(capellanes),
        "estados": len(estados),
    }


if __name__ == "__main__":
    print(extraer_catalogos_completos())
