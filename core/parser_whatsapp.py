import re
import unicodedata

from core.excel_store import leer_referencias


def normalizar(texto):
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFD", texto.lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9ñ\s]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _tokens(texto):
    return set(normalizar(texto).split())


def _buscar_catalogo(valor, opciones, preferidos=None):
    if not valor:
        return ""
    valor_norm = normalizar(valor)
    preferidos = preferidos or []

    alias = {
        "ctm": "Cambia Tu Mundo",
        "cambia tu mundo": "Cambia Tu Mundo",
        "rio verde": "Rio Verde",
        "impal": "Importadora Alemana",
        "importadora alemana": "Importadora Alemana",
    }
    objetivo = alias.get(valor_norm, valor)
    objetivo_norm = normalizar(objetivo)
    objetivo_tokens = _tokens(objetivo)

    for preferido in preferidos:
        for opcion in opciones:
            if normalizar(preferido) in normalizar(opcion):
                return opcion

    for opcion in opciones:
        opcion_norm = normalizar(opcion)
        if opcion_norm == objetivo_norm or objetivo_norm in opcion_norm or opcion_norm in objetivo_norm:
            return opcion

    for opcion in opciones:
        opcion_tokens = _tokens(opcion)
        if objetivo_tokens and objetivo_tokens.issubset(opcion_tokens):
            return opcion

    for opcion in opciones:
        opcion_tokens = _tokens(opcion)
        if opcion_tokens and opcion_tokens.issubset(objetivo_tokens):
            return opcion

    return objetivo


def _parsear_integrante(linea):
    linea = re.sub(r"^[\s\-\u2022\*\d\.\)]+", "", linea).strip()
    partes = [p.strip() for p in re.split(r"\s*[|,;]\s*", linea) if p.strip()]
    nombre_completo = partes[0] if partes else linea
    tokens = nombre_completo.split()
    nombre = tokens[0] if tokens else ""
    apellido = " ".join(tokens[1:]) if len(tokens) > 1 else ""

    data = {"Nombre*": nombre, "Apellido": apellido, "CI": "", "Email": "", "Sección": ""}
    for parte in partes[1:]:
        if "@" in parte:
            data["Email"] = parte
        elif re.search(r"\d", parte):
            data["CI"] = parte
        else:
            data["Sección"] = parte
    return data


def parsear_whatsapp(texto_bruto, catalogos=None):
    catalogos = catalogos or leer_referencias()
    lineas = [line.strip() for line in texto_bruto.splitlines() if line.strip()]
    cabecera = {}
    integrantes = []

    claves = {
        "grupo": "Grupo / Nombre*",
        "empresa": "Empresa*",
        "sucursal": "Sucursal*",
        "material": "Material*",
        "capellan": "Capellán*",
        "capellán": "Capellán*",
        "estado": "Estado Del Grupo*",
    }

    for line in lineas:
        es_integrante = bool(re.match(r"^(\u2022|-|\*|\d+[\.\)])\s+", line))
        if es_integrante:
            integrantes.append(_parsear_integrante(line))
            continue

        match = re.match(r"^([^:：=-]+)\s*[:：=-]\s*(.+)$", line)
        if match:
            clave = normalizar(match.group(1))
            destino = claves.get(clave)
            if destino:
                cabecera[destino] = match.group(2).strip()
            continue

        if "Grupo / Nombre*" not in cabecera:
            cabecera["Grupo / Nombre*"] = line
        else:
            integrantes.append(_parsear_integrante(line))

    empresa = _buscar_catalogo(
        cabecera.get("Empresa*", ""),
        catalogos.get("Empresas", []),
        preferidos=["Importadora Alemana", "HILAGRO S.A.E."] if normalizar(cabecera.get("Empresa*", "")) == "impal" else None,
    )
    sucursal = _buscar_catalogo(cabecera.get("Sucursal*", ""), catalogos.get("Sucursales", []))
    material = _buscar_catalogo(cabecera.get("Material*", ""), catalogos.get("Materiales", []))
    capellan = _buscar_catalogo(cabecera.get("Capellán*", ""), catalogos.get("Capellanes", []))
    estado = _buscar_catalogo(cabecera.get("Estado Del Grupo*", "ACTIVO"), catalogos.get("Estados", [])) or "ACTIVO"

    return [{
        "Grupo / Nombre*": cabecera.get("Grupo / Nombre*", "").strip(),
        "Empresa*": empresa,
        "Sucursal*": sucursal,
        "Capellán*": capellan,
        "Material*": material,
        "Estado Del Grupo*": estado,
        "Integrantes": integrantes,
    }] if cabecera.get("Grupo / Nombre*") else []
