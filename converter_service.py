import math
import os
import unicodedata
import openpyxl
import pandas as pd


def _normalizar_texto(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    texto = texto.strip().lower()
    nfkd = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])


def _normalizar_dcto(valor):
    """Convierte un valor de descuento a un entero o float representativo, o None si es inválido."""
    if pd.isna(valor) or valor == '' or valor is None:
        return None
    try:
        val_float = float(valor)
        if val_float.is_integer():
            return int(val_float)
        return val_float
    except (ValueError, TypeError):
        return str(valor).strip()


def _formatear_numero_negativo(val):
    """Convierte un valor a texto con signo negativo limpio (evita artefactos como 28.000000004)."""
    try:
        f = float(val)
        if f == 0.0:
            return '0'
        neg_f = -f
        if neg_f.is_integer():
            return str(int(neg_f))
        return f"{neg_f:.4f}".rstrip('0').rstrip('.')
    except (ValueError, TypeError):
        s = str(val).strip()
        if s.startswith('-'):
            return s
        return f"-{s}" if s else '0'


def cargar_datos_excel(ruta_excel, grupos_seleccionados):
    """Carga la hoja de Datos Iniciales y las tablas de descuento de los grupos seleccionados."""
    excel_file = pd.ExcelFile(ruta_excel, engine='openpyxl')
    nombres_hojas = excel_file.sheet_names
    hojas_norm_map = {_normalizar_texto(h): h for h in nombres_hojas}

    # 1. Buscar hoja 'Datos Iniciales'
    hoja_datos = None
    for k, original in hojas_norm_map.items():
        if k == 'datos iniciales' or 'datos inicial' in k:
            hoja_datos = original
            break

    if not hoja_datos:
        hojas_disp = ', '.join(nombres_hojas)
        raise ValueError(
            f"No se encontró la hoja 'Datos Iniciales'. Hojas disponibles: {hojas_disp}"
        )

    datos_iniciales = pd.read_excel(excel_file, sheet_name=hoja_datos)

    # 2. Buscar hojas de tablas de descuento
    tablas_descuento = {}
    hojas_faltantes = []

    for grupo in grupos_seleccionados:
        grupo_norm = _normalizar_texto(grupo)
        hoja_grupo = None

        for k, original in hojas_norm_map.items():
            if f"tabla descuentos {grupo_norm}" in k or f"descuentos {grupo_norm}" in k or k == f"t{grupo_norm}":
                hoja_grupo = original
                break

        if hoja_grupo:
            df_tabla = pd.read_excel(excel_file, sheet_name=hoja_grupo)
            tablas_descuento[grupo] = df_tabla
        else:
            hojas_faltantes.append(f"Tabla Descuentos {grupo}")

    if hojas_faltantes:
        hojas_disp = ', '.join(nombres_hojas)
        raise ValueError(
            f"Faltan las siguientes hojas de descuentos en el archivo Excel: {', '.join(hojas_faltantes)}. "
            f"Hojas encontradas: {hojas_disp}"
        )

    return datos_iniciales, tablas_descuento


def convertir_datos(datos_iniciales, tablas_descuento, valido_desde, num_acceso='35', cliente='', clientes=None):
    """Cruza los datos iniciales con las tablas de descuento de forma optimizada O(1).
    Soporta un cliente único o una lista de clientes (hasta N clientes).
    Todos los valores generados están estrictamente en formato Texto."""
    # Detectar nombres de columnas de 'Código' y 'DCTO'
    col_codigo = None
    col_dcto = None

    for col in datos_iniciales.columns:
        col_norm = _normalizar_texto(col)
        if col_norm in ['codigo', 'cod', 'material', 'cod_material']:
            col_codigo = col
        elif col_norm in ['dcto', 'descuento', 'dcto base']:
            col_dcto = col

    if not col_codigo:
        raise ValueError(
            f"No se encontró la columna de 'Código' en Datos Iniciales. Columnas encontradas: {list(datos_iniciales.columns)}"
        )
    if not col_dcto:
        raise ValueError(
            f"No se encontró la columna de 'DCTO' en Datos Iniciales. Columnas encontradas: {list(datos_iniciales.columns)}"
        )

    # Determinar lista de clientes
    if clientes is not None:
        lista_clientes = [str(c).strip() for c in clientes if c is not None and str(c).strip()]
        # Preservar orden único
        lista_clientes = list(dict.fromkeys(lista_clientes))
        if not lista_clientes:
            lista_clientes = ['']
    elif cliente is not None and str(cliente).strip():
        lista_clientes = [str(cliente).strip()]
    else:
        lista_clientes = ['']

    # Pre-indexar tablas de descuento en diccionarios para búsqueda instantánea O(1)
    dict_descuentos = {}
    for grupo, df in tablas_descuento.items():
        dict_descuentos[grupo] = {}

        col_t_dcto = None
        col_t_default = None
        for col in df.columns:
            cnorm = _normalizar_texto(col)
            if cnorm in ['dcto', 'descuento']:
                col_t_dcto = col
            elif cnorm in ['default', 'defecto', 'dcto default']:
                col_t_default = col

        if not col_t_dcto or not col_t_default:
            continue

        for idx, row in df.iterrows():
            d_norm = _normalizar_dcto(row[col_t_dcto])
            if d_norm is None:
                continue

            raw_def = row[col_t_default]
            if pd.isna(raw_def):
                continue
            imp_str = _formatear_numero_negativo(raw_def)

            servicios_dict = {'imp': imp_str}
            for tipo_servicio in ['R1', 'R2', 'R3', 'R4']:
                col_serv = None
                for c in df.columns:
                    if str(c).strip().upper() == tipo_servicio:
                        col_serv = c
                        break
                if col_serv and not pd.isna(row[col_serv]):
                    servicios_dict[tipo_servicio] = _formatear_numero_negativo(row[col_serv])

            dict_descuentos[grupo][d_norm] = servicios_dict

    columnas_resultado = [
        'ITEM',
        'Secuencia',
        '# acces',
        'Clase Cond',
        'Org. Vent',
        'Can Dist',
        'Sector',
        'Grupo Vend.',
        'Cliente Serv. (Preciario)',
        'Centro',
        'Grupo Clien',
        'Cliente',
        'Tipo Servicio',
        'Num Ext Veh',
        'Tipo Pral.VT',
        'Modelo',
        'Material',
        'Lista Prec',
        'Imp',
        'Un',
        'Lim Inf',
        'Lim Sup',
        'Valido de',
        'Valido a'
    ]

    lista_resultado = []
    item_counter = 1
    valido_desde_str = str(valido_desde).strip()
    acceso_str = str(num_acceso).strip() if num_acceso is not None else '35'

    codigos = datos_iniciales[col_codigo].values
    dctos = datos_iniciales[col_dcto].values

    for c_str in lista_clientes:
        for codigo_raw, dcto_raw in zip(codigos, dctos):
            if pd.isna(codigo_raw):
                continue
            codigo_str = str(codigo_raw).strip()
            if not codigo_str:
                continue

            dcto_norm = _normalizar_dcto(dcto_raw)
            if dcto_norm is None:
                continue

            for grupo, grupo_dict in dict_descuentos.items():
                descuentos = grupo_dict.get(dcto_norm)
                if not descuentos:
                    continue

                imp_str = descuentos['imp']

                for tipo_servicio in ['R1', 'R2', 'R3', 'R4']:
                    lim_inf_str = descuentos.get(tipo_servicio)
                    if lim_inf_str is None:
                        continue

                    fila = [
                        str(item_counter),
                        'Z2D1',
                        acceso_str,
                        'Z2D1',
                        '',
                        '',
                        '',
                        '',
                        '',
                        '',
                        str(grupo),
                        c_str,
                        '',
                        '',
                        '',
                        '',
                        codigo_str,
                        tipo_servicio,
                        imp_str,
                        '%',
                        lim_inf_str,
                        '0',
                        valido_desde_str,
                        '31.12.9999'
                    ]
                    lista_resultado.append(fila)
                    item_counter += 1

    return pd.DataFrame(lista_resultado, columns=columnas_resultado)


def exportar_en_lotes(df_resultado, carpeta_salida, prefijo_archivo, formato='csv', limite_filas=27000):
    """Divide el DataFrame en bloques de hasta `limite_filas` (por defecto 27,000),
    guardándolos en la carpeta indicada con formato Texto y encabezados repetidos en cada archivo.
    Retorna la lista de rutas de archivos generados."""
    os.makedirs(carpeta_salida, exist_ok=True)
    total_filas = len(df_resultado)

    if total_filas == 0:
        return []

    num_partes = math.ceil(total_filas / limite_filas)
    columnas = list(df_resultado.columns)
    archivos_generados = []

    es_xlsx = formato.lower() in ['xlsx', 'excel']
    extension = 'xlsx' if es_xlsx else 'csv'

    for i in range(num_partes):
        sub_df = df_resultado.iloc[i * limite_filas : (i + 1) * limite_filas]
        nombre_archivo = f"{prefijo_archivo}_Parte_{i + 1:02d}.{extension}"
        ruta_archivo = os.path.join(carpeta_salida, nombre_archivo)

        if es_xlsx:
            wb = openpyxl.Workbook(write_only=True)
            ws = wb.create_sheet()
            # Fila de encabezados
            ws.append(columnas)
            # Filas de datos (forzadas como texto)
            for row in sub_df.itertuples(index=False):
                ws.append([str(val) for val in row])
            wb.save(ruta_archivo)
        else:
            # CSV separado por comas, codificación UTF-8 con BOM
            sub_df.to_csv(ruta_archivo, index=False, sep=',', encoding='utf-8-sig')

        archivos_generados.append(ruta_archivo)

    return archivos_generados
