import streamlit as st
import pandas as pd
import pypdf
import re

st.set_page_config(page_title="Mi Conciliador Financiero", layout="wide")
st.title("📊 Conciliador Bancario y Tarjetas (Integración Fiwind)")

# 1. Función para leer el Excel de Banco Galicia (ARS y USD)
def leer_excel_galicia(archivo):
    try:
        df = pd.read_excel(archivo, skiprows=5)
        
        nombres_estandar = ["Fecha", "Movimiento", "Débito", "Crédito", "Saldo Parcial", "Comentarios"]
        columnas_actuales = df.columns.tolist()
        mapa_renombre = {columnas_actuales[i]: nombres_estandar[i] for i in range(min(len(columnas_actuales), len(nombres_estandar)))}
        df = df.rename(columns=mapa_renombre)
        
        if 'Fecha' not in df.columns:
            return pd.DataFrame()
            
        df = df.dropna(subset=['Fecha'])
        
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                df[col] = df[col].str.replace(r'[^\d.-]', '', regex=True)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                df[col] = 0.0
                
        df['Monto'] = df['Crédito'] + df['Débito']
        df['Categoria'] = 'Varios Bancario'
        
        df['Moneda'] = 'USD' if 'US' in archivo.name.upper() else 'ARS'
        
        df.loc[df['Movimiento'].str.contains('CUENTA PROPIA', case=False, na=False), 'Categoria'] = 'Transferencia Propia'
        df.loc[df['Movimiento'].str.contains('TITULOS|AL30', case=False, na=False), 'Categoria'] = 'Inversiones'
        df.loc[df['Movimiento'].str.contains('NORA', case=False, na=False), 'Categoria'] = 'Sueldo Nora'
        df.loc[df['Movimiento'].str.contains('ROCIO|ROCÍO', case=False, na=False), 'Categoria'] = 'Cuota Rocío'
        df.loc[df['Movimiento'].str.contains('ESTEBAN', case=False, na=False), 'Categoria'] = 'Celular Esteban'
        
        return df[['Fecha', 'Movimiento', 'Monto', 'Moneda', 'Categoria']]
    except Exception as e:
        st.error(f"Error procesando el archivo {archivo.name}.")
        return pd.DataFrame()

# 2. Función para extraer gastos de PDFs (Tarjetas y Mercado Pago)
def extraer_tarjeta_pdf(archivo, clave=""):
    data = []
    try:
        reader = pypdf.PdfReader(archivo)
        
        if reader.is_encrypted:
            # pypdf.PdfReader.decrypt devuelve 0 si la contraseña falla
            if reader.decrypt(clave) == 0:
                st.error(f"⚠️ Contraseña incorrecta para {archivo.name}.")
                return pd.DataFrame()

        for page in reader.pages:
            text = page.extract_text()
            if not text: continue
            
            for line in text.split('\n'):
                match = re.search(r'^(\d{2}/\d{2}/\d{2,4}|\d{2}-[A-Za-z]{3}-\d{2})\s+(.*?)\s+([\d\.\,]+)$', line.strip())
                if match:
                    fecha, desc, monto = match.groups()
                    
                    monto_limpio = monto.replace('$', '').replace(' ', '')
                    if ',' in monto_limpio and '.' in monto_limpio:
                        if monto_limpio.rfind(',') > monto_limpio.rfind('.'):
                            monto_limpio = monto_limpio.replace('.', '').replace(',', '.')
                        else:
                            monto_limpio = monto_limpio.replace(',', '')
                    elif ',' in monto_limpio:
                        monto_limpio = monto_limpio.replace(',', '.')
                    
                    try:
                        monto_float = -float(monto_limpio)
                    except ValueError:
                        continue
                        
                    data.append({
                        'Fecha': fecha,
                        'Movimiento': desc.strip(),
                        'Monto': monto_float,
                        'Moneda': 'ARS', 
                        'Categoria': 'Varios Tarjeta/MP'
                    })
                    
        df_tc = pd.DataFrame(data)
        
        if not df_tc.empty:
            df_tc.loc[df_tc['Movimiento'].str.contains('LIBERTAD|AL CAMPO|DISCO|SUPERMERCADO', case=False, na=False), 'Categoria'] = 'Supermercado'
            df_tc.loc[df_tc['Movimiento'].str.contains('FARMACITY|OMINT|SANATORIO', case=False, na=False), 'Categoria'] = 'Salud y Farmacia'
            df_tc.loc[df_tc['Movimiento'].str.contains('E.P.E.C|ECOGAS|MUNICIPALIDAD|CSIERRAS|NATURGY|EDESA|RENTAS', case=False, na=False), 'Categoria'] = 'Impuestos y Servicios'
            df_tc.loc[df_tc['Movimiento'].str.contains('YPF|AXION|COMBUSTIBLE', case=False, na=False), 'Categoria'] = 'Mantenimiento Vehículo'
            df_tc.loc[df_tc['Movimiento'].str.contains('UBER|PEDIDOS YA|PEDIDOSYA', case=False, na=False), 'Categoria'] = 'Apps y Transporte'

        return df_tc
        
    except Exception as e:
        st.error(f"⚠️ Hubo un error técnico al leer {archivo.name}: Verificá que el archivo no esté corrupto.")
        return pd.DataFrame()

# 3. Función para procesar operaciones Fiat en Fiwind
def leer_excel_fiwind(archivo):
    df = pd.read_excel(archivo)
    records = []
    
    for idx, row in df.iterrows():
        fecha_str = str(row['Fecha']).split(' ')[0]
        tipo = str(row['Tipo'])
        monto = float(row['Monto']) if pd.notna(row['Monto']) else 0.0
        moneda = str(row['Moneda'])
        monto_orig = float(row['Monto Origen']) if pd.notna(row['Monto Origen']) else 0.0
        moneda_orig = str(row['Moneda Origen'])
        
        if tipo == 'Conversión':
            if moneda in ['ARS', 'USD']:
                records.append({
                    'Fecha': fecha_str,
                    'Movimiento': f'Conversión desde {moneda_orig}',
                    'Monto': monto,
                    'Moneda': moneda,
                    'Categoria': 'Conversión Cripto/Fiwind'
                })
            if moneda_orig in ['ARS', 'USD']:
                records.append({
                    'Fecha': fecha_str,
                    'Movimiento': f'Conversión hacia {moneda}',
                    'Monto': -monto_orig,
                    'Moneda': moneda_orig,
                    'Categoria': 'Conversión Cripto/Fiwind'
                })
        else:
            if moneda in ['ARS', 'USD']:
                tipo_lower = tipo.lower()
                sign = -1 if 'retiro' in tipo_lower or 'pago' in tipo_lower else 1
                
                records.append({
                    'Fecha': fecha_str,
                    'Movimiento': tipo,
                    'Monto': monto * sign,
                    'Moneda': moneda,
                    'Categoria': 'Fiwind Fiat Varios'
                })
                
    df_fiwind = pd.DataFrame(records)
    
    if not df_fiwind.empty:
        df_fiwind.loc[df_fiwind['Movimiento'].str.contains('UBER|PEDIDOS', case=False, na=False), 'Categoria'] = 'Apps y Transporte'
        df_fiwind.loc[df_fiwind['Movimiento'].str.contains('Ganancia|Rendimiento', case=False, na=False), 'Categoria'] = 'Rendimiento Inversiones'
        df_fiwind.loc[df_fiwind['Movimiento'].str.contains('cuenta propia', case=False, na=False), 'Categoria'] = 'Transferencia Propia'
    
    return df_fiwind

# 4. Interfaz de la App
st.write("Sube tus archivos para comenzar la conciliación. El sistema consolidará todo en una única vista.")
mp_clave = st.text_input("🔑 Contraseña para archivos protegidos (Mercado Pago):", value="27549", type="password")

col1, col2, col3 = st.columns(3)
archivos_excel = col1.file_uploader("1️⃣ Excel Galicia", type=['xlsx'], accept_multiple_files=True)
archivos_pdf = col2.file_uploader("2️⃣ PDFs (Tarjetas/MP)", type=['pdf'], accept_multiple_files=True)
archivos_fiwind = col3.file_uploader("3️⃣ Excel Fiwind", type=['xlsx'], accept_multiple_files=True)

df_final = pd.DataFrame()

if archivos_excel:
    for excel in archivos_excel:
        df_final = pd.concat([df_final, leer_excel_galicia(excel)])

if archivos_pdf:
    for pdf in archivos_pdf:
        df_final = pd.concat([df_final, extraer_tarjeta_pdf(pdf, clave=mp_clave)])

if archivos_fiwind:
    for fiwind in archivos_fiwind:
        df_final = pd.concat([df_final, leer_excel_fiwind(fiwind)])

if not df_final.empty:
    st.success("¡Todos los datos fueron procesados y unificados con éxito!")
    
    monedas_disponibles = df_final['Moneda'].unique()
    
    col_a, col_b = st.columns([1, 2])
    with col_a:
        moneda_seleccionada = st.selectbox("Selecciona la moneda para visualizar:", monedas_disponibles)
        
    df_filtrado = df_final[df_final['Moneda'] == moneda_seleccionada].copy()
    
    st.write(f"### 📝 Registro Unificado de Movimientos ({moneda_seleccionada})")
    
    # Hemos quitado el atributo use_container_width para que Streamlit deje de enviar advertencias a la consola
    st.dataframe(df_filtrado)
    
    st.write(f"### 💰 Resumen de Gastos por Categoría ({moneda_seleccionada})")
    gastos = df_filtrado[df_filtrado['Monto'] < 0].copy()
    
    if not gastos.empty:
        gastos['Monto'] = gastos['Monto'].abs()
        resumen = gastos.groupby('Categoria')['Monto'].sum().reset_index()
        st.bar_chart(resumen, x='Categoria', y='Monto')
    else:
        st.info("No se registraron gastos (salidas de dinero) para esta moneda.")
    
    st.write("### ⬇️ Exportar CSV")
    csv = df_final.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Descargar tabla completa con todas las monedas",
        data=csv,
        file_name="movimientos_consolidados_fiat.csv",
        mime="text/csv",
    )
