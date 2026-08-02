#!/usr/bin/env python3
"""
Mi Conciliador Financiero — Versión mejorada
============================================
- Bug fixes: signos Galicia, parseo de montos AR, extracción PDF robusta, fechas
- Mejoras: normalización datetime, filtros, categorías configurables, Plotly
- Nuevas features: detección de duplicados, matching simple, reportes, Excel
- Refactor: modular, type hints, manejo de errores, session_state
"""

from __future__ import annotations

import re
import io
from datetime import datetime, date
from typing import Optional, List, Dict, Any, Tuple

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# PDF backend (prefer pypdf, fallback to PyPDF2)
# ---------------------------------------------------------------------------
try:
    from pypdf import PdfReader
    _PDF_BACKEND = "pypdf"
except ImportError:
    try:
        from PyPDF2 import PdfReader  # type: ignore
        _PDF_BACKEND = "PyPDF2"
    except ImportError:
        PdfReader = None  # type: ignore
        _PDF_BACKEND = None

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Mi Conciliador Financiero",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Configurable categories & rules
# ---------------------------------------------------------------------------
DEFAULT_CATEGORY_RULES: Dict[str, List[str]] = {
    "Transferencia Propia": [r"CUENTA PROPIA", r"cuenta propia"],
    "Inversiones": [r"TITULOS", r"AL30", r"ON\b", r"FCI", r"CEDEAR"],
    "Sueldo Nora": [r"NORA"],
    "Cuota Rocío": [r"ROCIO", r"ROCÍO"],
    "Celular Esteban": [r"ESTEBAN"],
    "Supermercado": [r"LIBERTAD", r"AL CAMPO", r"DISCO", r"SUPERMERCADO", r"CARREFOUR", r"JUMBO", r"VEA", r"DIA\b"],
    "Salud y Farmacia": [r"FARMACITY", r"OMINT", r"SANATORIO", r"FARMACIA", r"OSDE", r"SWISS MEDICAL"],
    "Impuestos y Servicios": [
        r"E\.P\.E\.C", r"ECOGAS", r"MUNICIPALIDAD", r"CSIERRAS", r"NATURGY",
        r"EDESA", r"RENTAS", r"AYSA", r"METROGAS", r"TELECOM", r"PERSONAL", r"CLARO", r"MOVISTAR",
    ],
    "Impuestos y Percepciones": [
        r"IMPUESTO\s+DE\s+SELLOS", r"IMP\.?\s*SELLOS", r"PERCEPCION", r"PERCEP\.?",
        r"RG\s*4815", r"DTO\.?\s*354", r"DTO\.?\s*2141", r"IMPUESTO\s+PAIS", r"IMP\.\s*PAIS",
    ],
    "Mantenimiento Ford Ranger 2013": [r"YPF", r"AXION", r"COMBUSTIBLE", r"SHELL", r"PUMASHELL"],
    "Apps y Transporte": [r"UBER", r"PEDIDOS YA", r"PEDIDOSYA", r"CABIFY", r"DIDI", r"RAPPI"],
    "Rendimiento Inversiones": [r"Ganancia", r"Rendimiento"],
    "Conversión Cripto/Fiwind": [r"Conversión"],
    "Fiwind Fiat Varios": [],  # fallback for Fiwind
    "Varios Bancario": [],
    "Varios Tarjeta/MP": [],
}


def apply_category_rules(df: pd.DataFrame, text_col: str = "Movimiento") -> pd.DataFrame:
    """Apply configurable regex rules. First match wins (order matters)."""
    if df.empty or text_col not in df.columns:
        return df
    df = df.copy()
    if "Categoria" not in df.columns:
        df["Categoria"] = "Sin categoría"

    for category, patterns in DEFAULT_CATEGORY_RULES.items():
        if not patterns:
            continue
        combined = "|".join(patterns)
        mask = df[text_col].str.contains(combined, case=False, na=False, regex=True)
        # Only overwrite if still default-ish
        default_mask = df["Categoria"].isin(
            ["Varios Bancario", "Varios Tarjeta/MP", "Fiwind Fiat Varios", "Sin categoría", ""]
        )
        df.loc[mask & default_mask, "Categoria"] = category
    return df


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------
def parse_argentine_amount(value: Any) -> Optional[float]:
    """Convert Argentine / mixed number formats to float. Returns None on failure."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip().replace("$", "").replace(" ", "").replace("\xa0", "")
    if not s or s.lower() in ("nan", "none", "-"):
        return None

    # 1.234.567,89  or  1.234,89  (AR)  vs  1,234.56 (US)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    # remove any remaining thousand separators that are dots only
    if s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(s)
    except ValueError:
        return None


def normalize_date(value: Any) -> Optional[pd.Timestamp]:
    """Parse common date formats found in AR bank / card statements."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime, pd.Timestamp, date)):
        return pd.Timestamp(value)

    s = str(value).strip().split(" ")[0]  # drop time if present
    formats = [
        "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y",
        "%d-%b-%Y", "%d-%b-%y", "%d-%B-%Y", "%d-%B-%y",
    ]
    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    # last resort
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return None


def safe_read_excel(uploaded_file, **kwargs) -> pd.DataFrame:
    """Read Excel with basic error handling."""
    try:
        return pd.read_excel(uploaded_file, **kwargs)
    except Exception as e:
        st.error(f"Error leyendo Excel '{getattr(uploaded_file, 'name', '?')}': {e}")
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# Source parsers
# ---------------------------------------------------------------------------
def leer_excel_galicia(archivo) -> pd.DataFrame:
    """Parse Banco Galicia extract (ARS or USD). Gastos (Débito) siempre negativos."""
    name = getattr(archivo, "name", str(archivo)).upper()
    df = safe_read_excel(
        archivo,
        skiprows=5,
        names=["Fecha", "Movimiento", "Débito", "Crédito", "Saldo Parcial", "Comentarios"],
    )
    if df.empty:
        return pd.DataFrame()

    df = df.dropna(subset=["Fecha"])
    if df.empty:
        return pd.DataFrame()

    df["Débito"] = df["Débito"].map(parse_argentine_amount)
    df["Crédito"] = df["Crédito"].map(parse_argentine_amount)

    # Gastos (Débito) negativos, créditos positivos
    debito = df["Débito"].fillna(0.0).abs()   # por si viene con signo
    credito = df["Crédito"].fillna(0.0).abs()
    df["Monto"] = credito - debito

    es_usd = "USD" in name
    df["Moneda"] = "USD" if es_usd else "ARS"
    df["Fuente"] = "Galicia USD" if es_usd else "Galicia ARS"
    df["Categoria"] = "Varios Bancario"
    df["Fecha"] = df["Fecha"].map(normalize_date)
    df["Comprobante"] = df["Comentarios"].fillna("").astype(str).str.strip()
    df.loc[df["Comprobante"].isin(["", "nan", "None"]), "Comprobante"] = ""

    df = apply_category_rules(df)
    return df[["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]].copy()


def _detect_card_fuente(filename: str, sample_text: str) -> str:
    """Detecta la tarjeta y devuelve el código de Fuente."""
    blob = f"{filename} {sample_text}".upper()
    # Orden importa: más específico primero
    if re.search(r"\bNX\s*VISA\b|\bNARANJA\s*X\s*VISA\b|\bTNX\s*VISA\b|\bVISA\s*NARANJA", blob):
        return "TNX Visa"
    if re.search(r"\bNARANJA\s*X\b|\bTARJETA\s*NARANJA\b|\bTNX\b|\bNARANJAX\b", blob):
        return "TNX"
    if re.search(r"\bMASTERCARD\b.*\bGALICIA\b|\bGALICIA\b.*\bMASTERCARD\b|\bTGM\b|\bMC\s*GALICIA\b", blob):
        return "TGM"
    if re.search(r"\bMASTERCARD\b|\bMASTER\s*CARD\b", blob):
        return "TGM"
    if re.search(r"\bVISA\b", blob):
        return "TNX Visa"
    if re.search(r"\bMERCADO\s*PAGO\b|\bMERCADOPAGO\b", blob):
        return "Mercado Pago"
    return "Tarjeta"


def _split_comprobante_from_desc(desc: str) -> Tuple[str, str]:
    """
    Separa el número de comprobante/cupón del final (o inicio) de la descripción.
    Ejemplos:
      'PAGOS360*MUNICORDO 04746'           → ('04746', 'PAGOS360*MUNICORDO')
      'FARMACITY INTERCOUNTRY 02/03 02382' → ('02382', 'FARMACITY INTERCOUNTRY 02/03')
      'Naranja X 54 OFICINA Y ARTE ...'    → ('54', 'OFICINA Y ARTE ...')
    """
    if not desc:
        return "", ""

    original = desc.strip()

    # 1) Número al final (4-8 dígitos), típico Galicia / TGM
    m = re.search(r"\s+(\d{4,8})$", original)
    if m:
        return m.group(1), original[: m.start()].strip()

    # 2) Patrón Naranja X / TNX: "Naranja X 54 DESCRIPCION" o "54 DESCRIPCION"
    m = re.match(r"^(?:NARANJA\s*X|TNX|NX)\s+(\d{2,8})\s+(.+)$", original, re.I)
    if m:
        return m.group(1), m.group(2).strip()

    # 3) Número corto al inicio seguido de texto
    m = re.match(r"^(\d{2,6})\s+(.+)$", original)
    if m:
        return m.group(1), m.group(2).strip()

    # 4) Etiquetas explícitas
    m = re.search(r"(?:CUP[OÓ]N|AUT|OPER|NRO|N[°º]|REF)[:\s#]*([0-9]{4,14})", original, re.I)
    if m:
        comp = m.group(1)
        cleaned = (original[: m.start()] + original[m.end() :]).strip()
        cleaned = re.sub(r"\s{2,}", " ", cleaned)
        return comp, cleaned

    return "", original


def extraer_tarjeta_pdf(archivo, clave: str = "") -> pd.DataFrame:
    """Extract transactions from credit-card / Mercado Pago PDFs.
    Robust to TNX / TGM / Visa layouts, taxes, perceptions, installments and previous balance.
    """
    if PdfReader is None:
        st.error("No hay backend PDF instalado (pypdf o PyPDF2).")
        return pd.DataFrame()

    data: List[Dict[str, Any]] = []
    filename = getattr(archivo, "name", str(archivo))
    try:
        reader = PdfReader(archivo)
        if getattr(reader, "is_encrypted", False):
            try:
                if reader.decrypt(clave or "") == 0:
                    st.error(f"⚠️ Contraseña incorrecta para {filename}.")
                    return pd.DataFrame()
            except Exception as e:
                st.error(f"⚠️ No se pudo desencriptar '{filename}'. Verifica la contraseña. ({e})")
                return pd.DataFrame()
    except Exception as e:
        st.error(f"Error abriendo PDF: {e}")
        return pd.DataFrame()

    sample_text = ""
    for page in reader.pages[:3]:
        try:
            sample_text += (page.extract_text() or "") + "\n"
        except Exception:
            pass
    fuente = _detect_card_fuente(filename, sample_text)

    def _parse_amount_neg(s: str) -> Optional[float]:
        v = parse_argentine_amount(s)
        return -v if v is not None else None

    def _es_consumo_usd(desc: str) -> bool:
        usd_merchants = [
            r"AMAZON\b", r"NETFLIX", r"SPOTIFY", r"GOOGLE", r"APPLE",
            r"STEAM", r"AIRBNB", r"BOOKING", r"ALIEXPRESS", r"PAYPAL",
            r"HBO", r"PRIME VIDEO", r"YOUTUBE", r"DISNEY", r"OPENAI", r"CHATGPT",
        ]
        return bool(re.search("|".join(usd_merchants), desc, re.I))

    # Líneas que deben ignorarse por completo
    SKIP_PAT = re.compile(
        r"SALDO\s+ANTERIOR|SALDO\s+ACTUAL|SALDO\s+CERRADO|SALDO\s+EN\s+PESOS|SALDO\s+EN\s+DOLARES|"
        r"SU\s+PAGO|PAGO\s+EN\s+PESOS|PAGO\s+EN\s+DOLARES|PAGO\s+AUTOMATICO|PAGO\s+DE\s+TARJETA|"
        r"TOTAL\s+A\s+PAGAR|TOTAL\s+CONSUMOS|SUBTOTAL|CONTIN[UÚ]A|HOJA\s+\d|"
        r"FECHA\s+DE\s+CIERRE|FECHA\s+DE\s+VENCIMIENTO|LIMITE\s+DE\s+COMPRA|"
        r"TASA\s+DE\s+INTERES|CFT|TEA|TEM|RESUMEN\s+DE\s+CUENTA|TITULAR|N[UÚ]MERO\s+DE\s+CUENTA",
        re.I,
    )

    # Impuestos / percepciones (TGM y otros)
    TAX_PAT = re.compile(
        r"IMPUESTO\s+DE\s+SELLOS|IMP\.?\s*SELLOS|PERCEPCION|PERCEP\.?|"
        r"RG\s*4815|DTO\.?\s*354|DTO\.?\s*2141|IVA\s+RG|IIBB|ING\.?\s*BRUTOS|"
        r"IMPUESTO\s+PAIS|IMP\.\s*PAIS|DEBITO\s+FISCAL|CREDITO\s+FISCAL",
        re.I,
    )

    DATE_PAT = re.compile(r"^(\d{2}/\d{2}/\d{2,4}|\d{2}-[A-Za-z]{3}-\d{2,4})$")
    # Monto solo en la línea (con o sin símbolo de moneda)
    AMT_ONLY_PAT = re.compile(
        r"^(?:(U\$S|USD|US\$|\$)\s*)?([\-\d\.\,]+)$", re.I
    )
    # Fecha + descripción + monto en la misma línea
    FULL_LINE_PAT = re.compile(
        r"^(\d{2}/\d{2}/\d{2,4}|\d{2}-[A-Za-z]{3}-\d{2,4})\s+(.*?)\s+(?:(U\$S|USD|US\$|\$)\s*)?([\-\d\.\,]+)$",
        re.I,
    )
    # Descripción + monto (sin fecha) — útil para cuotas e impuestos
    DESC_AMT_PAT = re.compile(
        r"^(.+?)\s+(?:(U\$S|USD|US\$|\$)\s*)?([\-\d\.\,]+)$", re.I
    )

    last_fecha: Optional[pd.Timestamp] = None

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            continue
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

        i = 0
        while i < len(lines):
            line = lines[i]

            # Ignorar encabezados / totales / saldo anterior
            if SKIP_PAT.search(line):
                i += 1
                continue

            # ---------- A) Línea completa: Fecha + Desc + Monto ----------
            m_full = FULL_LINE_PAT.match(line)
            if m_full:
                fecha_s, desc, curr_sym, monto_s = m_full.groups()
                if SKIP_PAT.search(desc):
                    i += 1
                    continue
                monto = _parse_amount_neg(monto_s)
                if monto is not None:
                    fecha = normalize_date(fecha_s)
                    if fecha is not None:
                        last_fecha = fecha
                    comp, clean_desc = _split_comprobante_from_desc(desc)
                    moneda = "USD" if (curr_sym and "U" in curr_sym.upper()) or _es_consumo_usd(clean_desc) else "ARS"
                    cat = "Impuestos y Percepciones" if TAX_PAT.search(clean_desc) else "Varios Tarjeta/MP"
                    data.append({
                        "Fecha": last_fecha,
                        "Comprobante": comp,
                        "Movimiento": clean_desc,
                        "Monto": monto,
                        "Moneda": moneda,
                        "Categoria": cat,
                        "Fuente": fuente,
                    })
                i += 1
                continue

            # ---------- B) Solo fecha → buscar desc + monto en siguientes líneas ----------
            if DATE_PAT.match(line):
                fecha_s = line
                fecha = normalize_date(fecha_s)
                if fecha is not None:
                    last_fecha = fecha
                desc_parts: List[str] = []
                monto = None
                moneda = "ARS"
                j = 1
                while j <= 8 and (i + j) < len(lines):
                    cand = lines[i + j]
                    if DATE_PAT.match(cand):
                        break
                    if SKIP_PAT.search(cand):
                        j += 1
                        continue
                    m_amt = AMT_ONLY_PAT.match(cand)
                    if m_amt:
                        curr_sym, raw_amt = m_amt.groups()
                        if curr_sym and "U" in curr_sym.upper():
                            moneda = "USD"
                        parsed = _parse_amount_neg(raw_amt)
                        if parsed is not None:
                            monto = parsed
                            j += 1
                            break
                    else:
                        if cand.upper() in ("U$S", "USD", "US$"):
                            moneda = "USD"
                        elif cand == "$":
                            moneda = "ARS"
                        else:
                            desc_parts.append(cand)
                    j += 1

                if monto is not None:
                    desc = " ".join(desc_parts).strip()
                    if not SKIP_PAT.search(desc):
                        comp, clean_desc = _split_comprobante_from_desc(desc)
                        if moneda == "ARS" and _es_consumo_usd(clean_desc):
                            moneda = "USD"
                        cat = "Impuestos y Percepciones" if TAX_PAT.search(clean_desc) else "Varios Tarjeta/MP"
                        data.append({
                            "Fecha": last_fecha,
                            "Comprobante": comp,
                            "Movimiento": clean_desc,
                            "Monto": monto,
                            "Moneda": moneda,
                            "Categoria": cat,
                            "Fuente": fuente,
                        })
                    i += j
                    continue
                i += 1
                continue

            # ---------- C) Sin fecha: impuestos, cuotas (03/03), o desc+monto ----------
            # Usa last_fecha si existe
            m_da = DESC_AMT_PAT.match(line)
            if m_da and last_fecha is not None:
                desc, curr_sym, monto_s = m_da.groups()
                # Evitar que una fecha sola o basura entre como descripción
                if DATE_PAT.match(desc.strip()):
                    i += 1
                    continue
                if SKIP_PAT.search(desc):
                    i += 1
                    continue
                # El "monto" no debe parecer una fecha (ej. 02/03)
                if re.match(r"^\d{2}/\d{2}$", monto_s.strip()):
                    i += 1
                    continue
                monto = _parse_amount_neg(monto_s)
                if monto is not None and abs(monto) >= 0.01:
                    comp, clean_desc = _split_comprobante_from_desc(desc)
                    moneda = "USD" if (curr_sym and "U" in curr_sym.upper()) or _es_consumo_usd(clean_desc) else "ARS"
                    is_tax = bool(TAX_PAT.search(clean_desc))
                    # Cuotas tipo "FARMACITY INTERCOUNTRY 03/03"
                    cat = "Impuestos y Percepciones" if is_tax else "Varios Tarjeta/MP"
                    data.append({
                        "Fecha": last_fecha,
                        "Comprobante": comp,
                        "Movimiento": clean_desc,
                        "Monto": monto,
                        "Moneda": moneda,
                        "Categoria": cat,
                        "Fuente": fuente,
                    })
                i += 1
                continue

            # ---------- D) Solo monto (raro) o línea de impuesto con monto al final ya cubierto ----------
            # Líneas tipo "IMPUESTO DE SELLOS" seguidas de monto en la próxima línea
            if TAX_PAT.search(line) and last_fecha is not None:
                desc = line
                monto = None
                moneda = "ARS"
                j = 1
                while j <= 3 and (i + j) < len(lines):
                    cand = lines[i + j]
                    if DATE_PAT.match(cand) or SKIP_PAT.search(cand):
                        break
                    m_amt = AMT_ONLY_PAT.match(cand)
                    if m_amt:
                        curr_sym, raw_amt = m_amt.groups()
                        if curr_sym and "U" in curr_sym.upper():
                            moneda = "USD"
                        parsed = _parse_amount_neg(raw_amt)
                        if parsed is not None:
                            monto = parsed
                            j += 1
                            break
                    j += 1
                if monto is not None:
                    data.append({
                        "Fecha": last_fecha,
                        "Comprobante": "",
                        "Movimiento": desc.strip(),
                        "Monto": monto,
                        "Moneda": moneda,
                        "Categoria": "Impuestos y Percepciones",
                        "Fuente": fuente,
                    })
                    i += j
                    continue

            i += 1

    df = pd.DataFrame(data)
    if df.empty:
        return df

    if "Comprobante" not in df.columns:
        df["Comprobante"] = ""
    # Limpiar comprobantes vacíos / basura
    df["Comprobante"] = df["Comprobante"].fillna("").astype(str)
    df = df.drop_duplicates(subset=["Fecha", "Movimiento", "Monto", "Fuente"])
    df = apply_category_rules(df)
    # Forzar categoría de impuestos si el nombre lo indica (por si apply_category_rules la pisó)
    tax_mask = df["Movimiento"].str.contains(
        r"IMPUESTO\s+DE\s+SELLOS|PERCEPCION|PERCEP\.?|RG\s*4815|DTO\.?\s*354|DTO\.?\s*2141|IMPUESTO\s+PAIS",
        case=False, na=False, regex=True,
    )
    df.loc[tax_mask, "Categoria"] = "Impuestos y Percepciones"
    return df[["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]].copy()


def leer_excel_fiwind(archivo) -> pd.DataFrame:
    """Process Fiwind fiat operations (incl. conversions)."""
    df = safe_read_excel(archivo)
    if df.empty:
        return pd.DataFrame()

    required = {"Fecha", "Tipo", "Monto", "Moneda"}
    if not required.issubset(set(df.columns)):
        st.warning(f"Fiwind: faltan columnas {required - set(df.columns)}. Se omite el archivo.")
        return pd.DataFrame()

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        fecha = normalize_date(row["Fecha"])
        tipo = str(row["Tipo"]).strip()
        monto = parse_argentine_amount(row["Monto"]) or 0.0
        moneda = str(row["Moneda"]).strip().upper()
        monto_orig = parse_argentine_amount(row.get("Monto Origen", 0)) or 0.0
        moneda_orig = str(row.get("Moneda Origen", "") or "").strip().upper()

        if tipo == "Conversión":
            if moneda in ("ARS", "USD"):
                records.append({
                    "Fecha": fecha,
                    "Comprobante": "",
                    "Movimiento": f"Conversión desde {moneda_orig or '?'}",
                    "Monto": monto,
                    "Moneda": moneda,
                    "Categoria": "Conversión Cripto/Fiwind",
                    "Fuente": "Fiwind",
                })
            if moneda_orig in ("ARS", "USD"):
                records.append({
                    "Fecha": fecha,
                    "Comprobante": "",
                    "Movimiento": f"Conversión hacia {moneda}",
                    "Monto": -monto_orig,
                    "Moneda": moneda_orig,
                    "Categoria": "Conversión Cripto/Fiwind",
                    "Fuente": "Fiwind",
                })
        else:
            if moneda in ("ARS", "USD"):
                tipo_lower = tipo.lower()
                sign = -1 if ("retiro" in tipo_lower or "pago" in tipo_lower) else 1
                records.append({
                    "Fecha": fecha,
                    "Comprobante": "",
                    "Movimiento": tipo,
                    "Monto": monto * sign,
                    "Moneda": moneda,
                    "Categoria": "Fiwind Fiat Varios",
                    "Fuente": "Fiwind",
                })

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        if "Comprobante" not in df_out.columns:
            df_out["Comprobante"] = ""
        df_out = apply_category_rules(df_out)
        df_out = df_out[["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]]
    return df_out


# ---------------------------------------------------------------------------
# Higher-level analytics
# ---------------------------------------------------------------------------
def detect_duplicates(df: pd.DataFrame, date_tol_days: int = 2, amount_tol: float = 1.0) -> pd.DataFrame:
    """Flag possible duplicates across sources (same |amount|, close dates)."""
    if df.empty or len(df) < 2:
        df = df.copy()
        df["Posible_Duplicado"] = False
        return df

    df = df.copy().reset_index(drop=True)
    df["Posible_Duplicado"] = False
    df["_abs"] = df["Monto"].abs()

    for i in range(len(df)):
        if df.at[i, "Posible_Duplicado"]:
            continue
        for j in range(i + 1, len(df)):
            if df.at[j, "Fuente"] == df.at[i, "Fuente"]:
                continue
            if abs(df.at[i, "_abs"] - df.at[j, "_abs"]) > amount_tol:
                continue
            d1, d2 = df.at[i, "Fecha"], df.at[j, "Fecha"]
            if pd.isna(d1) or pd.isna(d2):
                continue
            if abs((d1 - d2).days) <= date_tol_days:
                df.at[i, "Posible_Duplicado"] = True
                df.at[j, "Posible_Duplicado"] = True

    return df.drop(columns=["_abs"])


def build_monthly_report(df: pd.DataFrame) -> pd.DataFrame:
    """Monthly income / expense / net by currency."""
    if df.empty or "Fecha" not in df.columns:
        return pd.DataFrame()
    tmp = df.copy()
    tmp["Fecha"] = pd.to_datetime(tmp["Fecha"], errors="coerce")
    tmp = tmp.dropna(subset=["Fecha"])
    if tmp.empty:
        return pd.DataFrame()
    tmp["Mes"] = tmp["Fecha"].dt.to_period("M").astype(str)
    tmp["Tipo"] = tmp["Monto"].apply(lambda x: "Ingreso" if x > 0 else "Gasto")
    tmp["Monto_abs"] = tmp["Monto"].abs()

    pivot = (
        tmp.groupby(["Mes", "Moneda", "Tipo"])["Monto_abs"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )
    if "Ingreso" not in pivot.columns:
        pivot["Ingreso"] = 0.0
    if "Gasto" not in pivot.columns:
        pivot["Gasto"] = 0.0
    pivot["Neto"] = pivot["Ingreso"] - pivot["Gasto"]
    return pivot.sort_values(["Moneda", "Mes"])


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export DataFrame to formatted Excel bytes."""
    buffer = io.BytesIO()
    export_df = df.copy()
    if "Fecha" in export_df.columns:
        export_df["Fecha"] = pd.to_datetime(export_df["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Movimientos")
        wb = writer.book
        ws = writer.sheets["Movimientos"]
        header_fmt = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white"})
        for col_num, col_name in enumerate(export_df.columns):
            ws.write(0, col_num, col_name, header_fmt)
            try:
                max_len = max(export_df[col_name].astype(str).map(len).max(), len(str(col_name))) + 2
            except Exception:
                max_len = len(str(col_name)) + 2
            ws.set_column(col_num, col_num, min(int(max_len), 45))
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------
# Monedas que usan formato argentino: . miles  ,  decimal
_FIAT_AND_STABLE = {
    "ARS", "USD", "USDT", "USDC", "DAI", "BUSD", "TUSD", "USDP",
    "FRAX", "GUSD", "USDD", "FDUSD", "EURC", "PYUSD",
}


def format_monto_display(value: Any, moneda: str) -> str:
    """Formatea montos: ARS/USD/stablecoins con . miles y , decimal. Resto sin cambiar estilo."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value) if value is not None else ""

    mon = (moneda or "").upper()
    if mon in _FIAT_AND_STABLE:
        # Formato argentino: 1.234.567,89
        formatted = f"{v:,.2f}"  # 1,234,567.89
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
    # Otras criptos: formato estándar con punto decimal
    return f"{v:,.8f}".rstrip("0").rstrip(".") if abs(v) < 1 else f"{v:,.4f}".rstrip("0").rstrip(".")


def metric_row(df: pd.DataFrame, moneda: str):
    sub = df[df["Moneda"] == moneda]
    ingresos = sub.loc[sub["Monto"] > 0, "Monto"].sum()
    gastos = sub.loc[sub["Monto"] < 0, "Monto"].sum()  # already negative
    neto = sub["Monto"].sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Movimientos ({moneda})", f"{len(sub):,}")
    c2.metric("Ingresos", format_monto_display(ingresos, moneda))
    c3.metric("Gastos", format_monto_display(gastos, moneda))
    c4.metric("Neto", format_monto_display(neto, moneda), delta_color="normal")


# ---------------------------------------------------------------------------
# Main App
# ---------------------------------------------------------------------------
def main():
    st.title("📊 Conciliador Bancario y Tarjetas (Integración Fiwind)")
    st.caption(f"PDF backend: {_PDF_BACKEND or 'ninguno'} · Versión mejorada")

    # Sidebar – global options
    with st.sidebar:
        st.header("⚙️ Opciones")
        mp_clave = st.text_input(
            "🔑 Contraseña PDFs protegidos (ej. Mercado Pago)",
            value="27549",
            type="password",
        )
        st.markdown("---")
        st.subheader("Detección de duplicados")
        date_tol = st.slider("Tolerancia de fecha (días)", 0, 7, 2)
        amount_tol = st.number_input("Tolerancia de monto", min_value=0.0, value=1.0, step=0.5)
        st.markdown("---")
        st.subheader("Categorías")
        with st.expander("Ver / entender reglas"):
            for cat, pats in DEFAULT_CATEGORY_RULES.items():
                if pats:
                    st.markdown(f"**{cat}**")
                    st.code(" | ".join(pats), language="text")

    # File uploaders
    st.write("Sube tus archivos para comenzar la conciliación. El sistema unifica todo en una vista.")
    col1, col2, col3 = st.columns(3)
    with col1:
        archivos_excel = st.file_uploader(
            "1️⃣ Excel Galicia", type=["xlsx", "xls"], accept_multiple_files=True
        )
    with col2:
        archivos_pdf = st.file_uploader(
            "2️⃣ PDFs (Tarjetas / Mercado Pago)", type=["pdf"], accept_multiple_files=True
        )
    with col3:
        archivos_fiwind = st.file_uploader(
            "3️⃣ Excel Fiwind", type=["xlsx", "xls"], accept_multiple_files=True
        )

    # Botón centrado y con ancho ~1/3
    btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 1])
    with btn_col2:
        process = st.button("🔄 Procesar / Actualizar", type="primary", use_container_width=True)

    if process:
        frames: List[pd.DataFrame] = []

        if archivos_excel:
            for f in archivos_excel:
                with st.spinner(f"Galicia: {f.name}…"):
                    df = leer_excel_galicia(f)
                    if not df.empty:
                        frames.append(df)
                        st.success(f"✅ Galicia `{f.name}` → {len(df)} movimientos")
                    else:
                        st.warning(f"⚠️ Galicia `{f.name}` no produjo filas")

        if archivos_pdf:
            for f in archivos_pdf:
                with st.spinner(f"PDF: {f.name}…"):
                    df = extraer_tarjeta_pdf(f, clave=mp_clave)
                    if not df.empty:
                        frames.append(df)
                        st.success(f"✅ PDF `{f.name}` → {len(df)} movimientos")
                    else:
                        st.warning(f"⚠️ PDF `{f.name}` no produjo filas (¿layout distinto o vacío?)")

        if archivos_fiwind:
            for f in archivos_fiwind:
                with st.spinner(f"Fiwind: {f.name}…"):
                    df = leer_excel_fiwind(f)
                    if not df.empty:
                        frames.append(df)
                        st.success(f"✅ Fiwind `{f.name}` → {len(df)} movimientos")
                    else:
                        st.warning(f"⚠️ Fiwind `{f.name}` no produjo filas")

        if not frames:
            st.error("No se pudo extraer ningún movimiento. Revisá los archivos e intentá de nuevo.")
            st.stop()

        df_final = pd.concat(frames, ignore_index=True)
        if "Comprobante" not in df_final.columns:
            df_final["Comprobante"] = ""
        df_final["Comprobante"] = df_final["Comprobante"].fillna("").astype(str)
        # Ensure Fecha is datetime
        df_final["Fecha"] = pd.to_datetime(df_final["Fecha"], errors="coerce")
        # Orden de columnas: Fecha, Comprobante, ...
        cols = ["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]
        extra = [c for c in df_final.columns if c not in cols]
        df_final = df_final[[c for c in cols if c in df_final.columns] + extra]
        df_final = detect_duplicates(df_final, date_tol_days=date_tol, amount_tol=amount_tol)
        df_final = df_final.sort_values("Fecha", na_position="last").reset_index(drop=True)
        st.session_state.df_final = df_final
        st.success(f"🎉 Procesados y unificados **{len(df_final)}** movimientos.")

    if "df_final" not in st.session_state or st.session_state.df_final is None:
        st.info("👆 Subí al menos un archivo y presioná **Procesar / Actualizar** para comenzar.")
        return

    df_final: pd.DataFrame = st.session_state.df_final.copy()
    if "Fecha" in df_final.columns:
        df_final["Fecha"] = pd.to_datetime(df_final["Fecha"], errors="coerce")

    # ------------------ Filters ------------------
    st.markdown("---")
    st.subheader("🔍 Filtros")
    f1, f2, f3, f4 = st.columns(4)

    monedas = sorted(df_final["Moneda"].dropna().unique().tolist())
    with f1:
        moneda_sel = st.multiselect("Moneda", monedas, default=monedas)
    with f2:
        cats = sorted(df_final["Categoria"].dropna().unique().tolist())
        cat_sel = st.multiselect("Categoría", cats, default=cats)
    with f3:
        fuentes = sorted(df_final["Fuente"].dropna().unique().tolist())
        fuente_sel = st.multiselect("Fuente", fuentes, default=fuentes)
    with f4:
        only_dups = st.checkbox("Solo posibles duplicados", value=False)

    # Date range
    start_d = end_d = None
    valid_dates = df_final["Fecha"].dropna()
    if not valid_dates.empty:
        min_d, max_d = valid_dates.min().date(), valid_dates.max().date()
        dr = st.date_input(
            "Rango de fechas",
            value=(min_d, max_d),
            min_value=min_d,
            max_value=max_d,
        )
        if isinstance(dr, (list, tuple)) and len(dr) == 2:
            start_d, end_d = dr
        else:
            start_d = end_d = dr if not isinstance(dr, (list, tuple)) else min_d

    mask = (
        df_final["Moneda"].isin(moneda_sel)
        & df_final["Categoria"].isin(cat_sel)
        & df_final["Fuente"].isin(fuente_sel)
    )
    if start_d is not None and end_d is not None:
        mask &= df_final["Fecha"].isna() | (
            (df_final["Fecha"] >= pd.Timestamp(start_d))
            & (df_final["Fecha"] <= pd.Timestamp(end_d))
        )
    if only_dups:
        mask &= df_final["Posible_Duplicado"].fillna(False)

    df_filt = df_final.loc[mask].copy()

    # ------------------ Tabs ------------------
    tab_mov, tab_res, tab_rep, tab_dup = st.tabs(
        ["📝 Movimientos", "💰 Resumen & Charts", "📅 Reportes mensuales", "🔁 Posibles duplicados"]
    )

    with tab_mov:
        for m in moneda_sel:
            st.markdown(f"#### {m}")
            metric_row(df_filt, m)

        # Display copy: fechas legibles + montos con formato AR para fiat/stablecoins
        display_df = df_filt.copy()
        if "Fecha" in display_df.columns:
            display_df["Fecha"] = display_df["Fecha"].dt.strftime("%Y-%m-%d").fillna("")
        if "Comprobante" not in display_df.columns:
            display_df["Comprobante"] = ""
        # Formatear Monto según moneda (ARS/USD/stable → 1.234,56 ; otras cripto sin cambiar)
        if "Monto" in display_df.columns and "Moneda" in display_df.columns:
            display_df["Monto"] = [
                format_monto_display(v, m)
                for v, m in zip(display_df["Monto"], display_df["Moneda"])
            ]
        # Orden: Fecha | Comprobante | resto
        preferred = ["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente", "Posible_Duplicado"]
        ordered = [c for c in preferred if c in display_df.columns]
        ordered += [c for c in display_df.columns if c not in ordered]
        display_df = display_df[ordered]
        st.dataframe(display_df, use_container_width=True, height=420)

        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            csv = df_final.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Descargar CSV completo",
                data=csv,
                file_name="movimientos_consolidados.csv",
                mime="text/csv",
            )
        with c_dl2:
            try:
                xlsx_bytes = to_excel_bytes(df_final)
                st.download_button(
                    "⬇️ Descargar Excel formateado",
                    data=xlsx_bytes,
                    file_name="movimientos_consolidados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.caption(f"Excel no disponible ({e})")

    with tab_res:
        if df_filt.empty:
            st.info("No hay datos con los filtros actuales.")
        else:
            for moneda in moneda_sel:
                st.markdown(f"### {moneda}")
                sub = df_filt[df_filt["Moneda"] == moneda]
                gastos = sub[sub["Monto"] < 0].copy()
                if gastos.empty:
                    st.write("Sin gastos registrados.")
                    continue
                gastos["Monto_abs"] = gastos["Monto"].abs()
                resumen = (
                    gastos.groupby("Categoria", as_index=False)["Monto_abs"]
                    .sum()
                    .sort_values("Monto_abs", ascending=False)
                )

                c1, c2 = st.columns([1, 1])
                with c1:
                    fig = px.bar(
                        resumen,
                        x="Categoria",
                        y="Monto_abs",
                        title=f"Gastos por categoría ({moneda})",
                        labels={"Monto_abs": "Monto", "Categoria": ""},
                        text_auto=".2s",
                    )
                    fig.update_layout(xaxis_tickangle=-35, height=400)
                    st.plotly_chart(fig, use_container_width=True)
                with c2:
                    fig2 = px.pie(
                        resumen,
                        names="Categoria",
                        values="Monto_abs",
                        title=f"Distribución de gastos ({moneda})",
                        hole=0.35,
                    )
                    fig2.update_layout(height=400)
                    st.plotly_chart(fig2, use_container_width=True)

                resumen_disp = resumen.copy()
                resumen_disp["Monto_abs"] = resumen_disp["Monto_abs"].map(
                    lambda x: format_monto_display(x, moneda)
                )
                st.dataframe(resumen_disp, use_container_width=True)

    with tab_rep:
        report = build_monthly_report(df_filt)
        if report.empty:
            st.info("No hay datos suficientes para el reporte mensual.")
        else:
            report_disp = report.copy()
            for col in ("Ingreso", "Gasto", "Neto"):
                if col in report_disp.columns:
                    report_disp[col] = [
                        format_monto_display(v, m)
                        for v, m in zip(report_disp[col], report_disp["Moneda"])
                    ]
            st.dataframe(report_disp, use_container_width=True)
            for moneda in moneda_sel:
                r = report[report["Moneda"] == moneda]
                if r.empty:
                    continue
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Ingresos", x=r["Mes"], y=r["Ingreso"], marker_color="#2ecc71"))
                fig.add_trace(go.Bar(name="Gastos", x=r["Mes"], y=r["Gasto"], marker_color="#e74c3c"))
                fig.add_trace(go.Scatter(
                    name="Neto", x=r["Mes"], y=r["Neto"], mode="lines+markers",
                    line=dict(color="#3498db", width=3),
                ))
                fig.update_layout(
                    barmode="group",
                    title=f"Evolución mensual ({moneda})",
                    height=420,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig, use_container_width=True)

    with tab_dup:
        dups = df_final[df_final["Posible_Duplicado"].fillna(False)]
        st.write(
            f"Se detectaron **{len(dups)}** movimientos marcados como posibles duplicados "
            f"(tolerancia ±{date_tol} días / ±{amount_tol} de monto, distintas fuentes)."
        )
        if dups.empty:
            st.success("No se encontraron posibles duplicados con los parámetros actuales.")
        else:
            disp = dups.copy()
            if "Fecha" in disp.columns:
                disp["Fecha"] = disp["Fecha"].dt.strftime("%Y-%m-%d").fillna("")
            if "Comprobante" not in disp.columns:
                disp["Comprobante"] = ""
            if "Monto" in disp.columns and "Moneda" in disp.columns:
                disp["Monto"] = [
                    format_monto_display(v, m)
                    for v, m in zip(disp["Monto"], disp["Moneda"])
                ]
            preferred = ["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]
            ordered = [c for c in preferred if c in disp.columns]
            ordered += [c for c in disp.columns if c not in ordered]
            disp = disp[ordered]
            st.dataframe(disp, use_container_width=True, height=400)
            st.info(
                "Revisá manualmente: un mismo gasto puede aparecer en el extracto bancario "
                "y también en el resumen de la tarjeta o en Fiwind."
            )


if __name__ == "__main__":
    main()