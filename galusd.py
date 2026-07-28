# 1. Función para leer el Excel de Banco Galicia (ARS y USD)
def leer_excel_galicia(archivo):
    try:
        # Leemos el archivo sin forzar nombres fijos para evitar errores de cantidad de columnas
        df = pd.read_excel(archivo, skiprows=5)
        
        # Renombramos las primeras columnas basándonos en el estándar de Galicia
        nombres_estandar = ["Fecha", "Movimiento", "Débito", "Crédito", "Saldo Parcial", "Comentarios"]
        columnas_actuales = df.columns.tolist()
        mapa_renombre = {columnas_actuales[i]: nombres_estandar[i] for i in range(min(len(columnas_actuales), len(nombres_estandar)))}
        df = df.rename(columns=mapa_renombre)
        
        # Si el archivo no tiene la estructura mínima, devolvemos un DataFrame vacío
        if 'Fecha' not in df.columns:
            return pd.DataFrame()
            
        df = df.dropna(subset=['Fecha'])
        
        # Limpieza robusta de montos (Débito y Crédito)
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                # Convertimos a texto, quitamos puntos de miles y cambiamos comas a puntos
                df[col] = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                # Eliminamos cualquier letra o símbolo que no sea un número, punto o guion
                df[col] = df[col].str.replace(r'[^\d.-]', '', regex=True)
                # Convertimos a número. Si algo falla (errors='coerce'), lo vuelve 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                # Si la columna no existe en este archivo, la creamos en 0
                df[col] = 0.0
                
        df['Monto'] = df['Crédito'] + df['Débito']
        df['Categoria'] = 'Varios Bancario'
        
        # Detectar moneda desde el nombre del archivo
        df['Moneda'] = 'USD' if 'US' in archivo.name.upper() else 'ARS'
        
        # Reglas automáticas para movimientos bancarios
        df.loc[df['Movimiento'].str.contains('CUENTA PROPIA', case=False, na=False), 'Categoria'] = 'Transferencia Propia'
        df.loc[df['Movimiento'].str.contains('TITULOS|AL30', case=False, na=False), 'Categoria'] = 'Inversiones'
        df.loc[df['Movimiento'].str.contains('NORA', case=False, na=False), 'Categoria'] = 'Sueldo Nora'
        df.loc[df['Movimiento'].str.contains('ROCIO|ROCÍO', case=False, na=False), 'Categoria'] = 'Cuota Rocío'
        df.loc[df['Movimiento'].str.contains('ESTEBAN', case=False, na=False), 'Categoria'] = 'Celular Esteban'
        
        return df[['Fecha', 'Movimiento', 'Monto', 'Moneda', 'Categoria']]
        
    except Exception as e:
        # En caso de error, mostramos un aviso en la web en lugar de romper toda la app
        st.error(f"Error procesando el archivo {archivo.name}. Se omitirá en el consolidado.")
        return pd.DataFrame()