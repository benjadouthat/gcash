def extraer_tarjeta_pdf(archivo, clave: str = "") -> pd.DataFrame:
    """Extract transactions from credit-card / Mercado Pago PDFs.
    Handles both single-line and multi-line layouts. Fixes taxes, USD, and payments.
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
                # pypdf devuelve 0 si la contraseña falla
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

    # Heurística: si el banco no pone U$S, detectamos comercios típicos en dólares
    def _es_consumo_usd(desc: str) -> bool:
        usd_merchants = [
            r"AMAZON\b", r"NETFLIX", r"SPOTIFY", r"GOOGLE", r"APPLE", 
            r"STEAM", r"AIRBNB", r"BOOKING", r"ALIEXPRESS", r"PAYPAL", 
            r"HBO", r"PRIME VIDEO", r"YOUTUBE", r"DISNEY"
        ]
        return bool(re.search("|".join(usd_merchants), desc, re.I))

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            continue
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

        i = 0
        while i < len(lines):
            line = lines[i]

            # --- Estrategia A: Todo en una sola línea (Regex Clásica) ---
            m_single = re.search(
                r"^(\d{2}/\d{2}/\d{2,4}|\d{2}-[A-Za-z]{3}-\d{2,4})\s+(.*?)\s+(?:(U\$S|USD|US\$|\$)\s*)?([\-\d\.\,]+)$",
                line, re.I
            )
            if m_single:
                fecha_s, desc, curr_sym, monto_s = m_single.groups()
                
                # Omitimos el pago del resumen para no alterar el neto
                if re.search(r"SU PAGO|PAGO EN PESOS|PAGO EN DOLARES|PAGO AUTOMATICO|PAGO DE TARJETA", desc, re.I):
                    i += 1
                    continue

                monto = _parse_amount_neg(monto_s)
                if monto is not None:
                    comp, clean_desc = _split_comprobante_from_desc(desc)
                    # Asignación inteligente de USD
                    moneda = "USD" if (curr_sym and "U" in curr_sym.upper()) or _es_consumo_usd(clean_desc) else "ARS"
                    
                    data.append({
                        "Fecha": normalize_date(fecha_s),
                        "Comprobante": comp,
                        "Movimiento": clean_desc,
                        "Monto": monto,
                        "Moneda": moneda,
                        "Categoria": "Varios Tarjeta/MP",
                        "Fuente": fuente,
                    })
                i += 1
                continue

            # --- Estrategia B: Multilínea (Fecha -> Desc -> Monto) ---
            if re.match(r"^\d{2}/\d{2}/\d{2,4}$", line) or re.match(r"^\d{2}-[A-Za-z]{3}-\d{2,4}$", line):
                fecha_s = line
                desc_parts = []
                monto = None
                moneda = "ARS"
                
                j = 1
                # Miramos hasta 6 líneas hacia adelante (ideal para capturar impuestos dispersos)
                while j <= 6 and (i + j) < len(lines):
                    cand = lines[i + j]
                    
                    # Si chocamos con otra fecha, detenemos la búsqueda
                    if re.match(r"^\d{2}/\d{2}/\d{2,4}$", cand) or re.match(r"^\d{2}-[A-Za-z]{3}-\d{2,4}$", cand):
                        break
                        
                    # Buscamos el monto (puede incluir signos negativos o de moneda)
                    m_amt = re.match(r"^(?:(U\$S|USD|US\$|\$)\s*)?([\-\d\.\,]+)$", cand, re.I)
                    if m_amt:
                        curr_sym = m_amt.group(1)
                        if curr_sym and "U" in curr_sym.upper():
                            moneda = "USD"
                        
                        monto = _parse_amount_neg(m_amt.group(2))
                        if monto is not None:
                            j += 1 # Consumimos la línea del monto
                            break
                    else:
                        # Atrapamos símbolos de moneda sueltos en un renglón
                        if cand.upper() in ["U$S", "USD", "US$"]:
                            moneda = "USD"
                        elif cand == "$":
                            moneda = "ARS"
                        else:
                            desc_parts.append(cand)
                    j += 1
                
                if monto is not None:
                    desc = " ".join(desc_parts).strip()
                    
                    # Omitimos el pago del resumen
                    if re.search(r"SU PAGO|PAGO EN PESOS|PAGO EN DOLARES|PAGO AUTOMATICO|PAGO DE TARJETA", desc, re.I):
                        i += j
                        continue

                    comp, clean_desc = _split_comprobante_from_desc(desc)
                    
                    # Verificación final de comercios USD por si faltó el símbolo
                    if moneda == "ARS" and _es_consumo_usd(clean_desc):
                        moneda = "USD"

                    data.append({
                        "Fecha": normalize_date(fecha_s),
                        "Comprobante": comp,
                        "Movimiento": clean_desc,
                        "Monto": monto,
                        "Moneda": moneda,
                        "Categoria": "Varios Tarjeta/MP",
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
    df = df.drop_duplicates(subset=["Fecha", "Movimiento", "Monto"])
    df = apply_category_rules(df)
    return df[["Fecha", "Comprobante", "Movimiento", "Monto", "Moneda", "Categoria", "Fuente"]].copy()