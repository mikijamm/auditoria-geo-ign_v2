# -*- coding: utf-8 -*-
"""
Auditoría Geoespacial y Validación Municipal de Catálogos (IGN)
Aplicación Streamlit basada en Auditoria_GEO.ipynb.
"""

from __future__ import annotations

import difflib
import html
import io
import re
import unicodedata
from typing import Any, Dict, List, Optional, Sequence, Tuple

import folium
import geopandas as gpd
import numpy as np
import pandas as pd
import shapely
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pyproj import Transformer

# ==============================================================================
# CONSTANTES Y CONFIGURACIÓN GLOBAL
# ==============================================================================

URL_IGN_MUNICIPIOS = (
    "https://github.com/rOpenSpain/mapSpain/raw/sianedata/dist/"
    "se89_3_admin_muni_a_x.gpkg"
)
URL_IGN_MUNICIPIOS_CANARIAS = (
    "https://github.com/rOpenSpain/mapSpain/raw/sianedata/dist/"
    "se89_3_admin_muni_a_y.gpkg"
)

PROVINCIAS_INE: Dict[str, str] = {
    "01": "Álava",
    "02": "Albacete",
    "03": "Alicante",
    "04": "Almería",
    "05": "Ávila",
    "06": "Badajoz",
    "07": "Illes Balears",
    "08": "Barcelona",
    "09": "Burgos",
    "10": "Cáceres",
    "11": "Cádiz",
    "12": "Castellón",
    "13": "Ciudad Real",
    "14": "Córdoba",
    "15": "A Coruña",
    "16": "Cuenca",
    "17": "Girona",
    "18": "Granada",
    "19": "Guadalajara",
    "20": "Gipuzkoa",
    "21": "Huelva",
    "22": "Huesca",
    "23": "Jaén",
    "24": "León",
    "25": "Lleida",
    "26": "La Rioja",
    "27": "Lugo",
    "28": "Madrid",
    "29": "Málaga",
    "30": "Murcia",
    "31": "Navarra",
    "32": "Ourense",
    "33": "Asturias",
    "34": "Palencia",
    "35": "Las Palmas",
    "36": "Pontevedra",
    "37": "Salamanca",
    "38": "Santa Cruz de Tenerife",
    "39": "Cantabria",
    "40": "Segovia",
    "41": "Sevilla",
    "42": "Soria",
    "43": "Tarragona",
    "44": "Teruel",
    "45": "Toledo",
    "46": "Valencia",
    "47": "Valladolid",
    "48": "Bizkaia",
    "49": "Zamora",
    "50": "Zaragoza",
    "51": "Ceuta",
    "52": "Melilla",
}

PROVINCIAS_CANARIAS = {"35", "38"}

PRIORIDAD_ESTADOS: List[str] = [
    "ERROR_GRAVE",
    "FUERA_DE_PROVINCIA_O_SIN_POLIGONO",
    "DESVÍO_NOTABLE",
    "DESVÍO_MODERADO",
    "BORDE_CERCANO",
    "FRONTERIZO_INMEDIATO",
    "TOPÓNIMO_NO_LOCALIZADO",
    "CORRECTO",
]

ESTILOS_MAPA: Dict[str, Dict[str, Any]] = {
    "ERROR_GRAVE": {"color": "#922b21", "fill": "#e74c3c", "r": 8.0},
    "FUERA_DE_PROVINCIA_O_SIN_POLIGONO": {
        "color": "#5b2c6f",
        "fill": "#9b59b6",
        "r": 8.5,
    },
    "DESVÍO_NOTABLE": {"color": "#6e2c00", "fill": "#d35400", "r": 7.0},
    "DESVÍO_MODERADO": {"color": "#b9770e", "fill": "#f39c12", "r": 6.0},
    "BORDE_CERCANO": {"color": "#b7950b", "fill": "#f1c40f", "r": 5.0},
    "FRONTERIZO_INMEDIATO": {"color": "#196f3d", "fill": "#2ecc71", "r": 4.5},
    "TOPÓNIMO_NO_LOCALIZADO": {"color": "#566573", "fill": "#aab7b8", "r": 5.0},
    "CORRECTO": {"color": "#1f618d", "fill": "#3498db", "r": 4.0},
}

ARTICULOS_INVERTIDOS = {
    "el",
    "la",
    "los",
    "las",
    "lo",
    "els",
    "les",
    "ses",
    "es",
    "sa",
    "o",
    "a",
    "os",
    "as",
    "l",
    "s",
}

SINONIMOS_COLUMNAS: Dict[str, List[str]] = {
    "id": [
        "id",
        "codigo",
        "cod",
        "n",
        "no",
        "num",
        "numero",
        "clave",
        "identificador",
        "ref",
        "referencia",
        "expediente",
        "registro",
        "id recurso",
        "codigo recurso",
        "id activo",
        "code",
    ],
    "recurso": [
        "nombre establecimiento",
        "recurso",
        "denominacion",
        "nombre",
        "activo",
        "nombre activo",
        "nombre recurso",
        "establecimiento",
        "titulo",
        "entidad",
        "instalacion",
        "equipamiento",
        "empresa",
        "centro",
        "nombre comercial",
        "razon social",
        "lugar",
        "sitio",
        "name",
        "title",
    ],
    "municipio": [
        "municipio",
        "poblacion",
        "localidad",
        "pueblo",
        "ciudad",
        "ayuntamiento",
        "termino municipal",
        "nombre municipio",
        "muni",
        "mun",
        "concejo",
        "municipality",
        "town",
        "city",
    ],
    "latitud": [
        "latitud",
        "lat",
        "y",
        "coord y",
        "coordenada y",
        "latitude",
        "gps lat",
        "lat wgs84",
        "latitud wgs84",
        "latitud gps",
        "geo lat",
        "lat y",
        "y lat",
    ],
    "longitud": [
        "longitud",
        "lon",
        "lng",
        "long",
        "x",
        "coord x",
        "coordenada x",
        "longitude",
        "gps lon",
        "lon wgs84",
        "longitud wgs84",
        "longitud gps",
        "geo lon",
        "lon x",
        "x lon",
    ],
    "descripcion": [
        "descripcion corta",
        "descripcion",
        "resumen",
        "detalle",
        "observaciones",
        "notas",
        "description",
    ],
    "direccion": [
        "direccion",
        "domicilio",
        "calle",
        "via",
        "emplazamiento",
        "ubicacion",
        "address",
        "dir",
    ],
    "url": [
        "url fuente",
        "url",
        "web",
        "fuente",
        "enlace",
        "link",
        "pagina web",
        "sitio web",
        "website",
    ],
}

SIN_COLUMNA_OPCION = "— Ninguna / No disponible —"

# Transformadores pyproj reutilizables para distancias métricas precisas
_TRANS_4326_TO_25830 = Transformer.from_crs("EPSG:4326", "EPSG:25830", always_xy=True)
_TRANS_4326_TO_25828 = Transformer.from_crs("EPSG:4326", "EPSG:25828", always_xy=True)


# ==============================================================================
# FUNCIONES PURAS DE NORMALIZACIÓN Y LIMPIEZA
# ==============================================================================


def normalizar_texto(texto: Any) -> str:
    """Normaliza texto: minúsculas, sin tildes y espacios colapsados."""
    if texto is None or (isinstance(texto, float) and np.isnan(texto)):
        return ""
    try:
        if pd.isna(texto):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(texto).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalizar_nombre_columna(nombre: Any) -> str:
    """Normaliza un nombre de columna para comparaciones flexibles."""
    base = normalizar_texto(nombre)
    base = re.sub(r"[_\-./()\[\]:;ºª°]+", " ", base)
    base = re.sub(r"\s+", " ", base)
    return base.strip()


def normalizar_toponimo(texto: Any) -> str:
    """
    Normaliza un topónimo español resolviendo artículos invertidos entre paréntesis
    o tras coma (ej. 'ALDEA DEL OBISPO (LA)' -> 'la aldea del obispo',
    'TORNO, EL' -> 'el torno', 'HOSPITALET (L')' -> 'l hospitalet').
    """
    s = normalizar_texto(texto)
    if not s:
        return ""

    # 1. Artículo invertido entre paréntesis al final: "aldea del obispo (la)"
    m_par = re.match(r"^(.+?)\s*\(([^()]+)\)\s*$", s)
    if m_par:
        cuerpo = m_par.group(1).strip()
        dentro = re.sub(r"['’`]", "", m_par.group(2)).strip()
        if dentro in ARTICULOS_INVERTIDOS:
            s = f"{dentro} {cuerpo}"

    # 2. Artículo invertido tras coma al final: "aldea del obispo, la"
    m_coma = re.match(r"^(.+?)\s*,\s*([^,]+)\s*$", s)
    if m_coma:
        cuerpo = m_coma.group(1).strip()
        tras_coma = re.sub(r"['’`]", "", m_coma.group(2)).strip()
        if tras_coma in ARTICULOS_INVERTIDOS:
            s = f"{tras_coma} {cuerpo}"

    # 3. Limpiar puntuación menor (apóstrofos, guiones) preservando separador '/'
    s = re.sub(r"['’`]", " ", s)
    s = re.sub(r"\s*-\s*", " ", s)
    s = re.sub(r"[.,;:!?]", " ", s)
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def variantes_toponimo(texto: Any) -> List[str]:
    """
    Devuelve variantes normalizadas de un topónimo, contemplando nombres
    cooficiales separados por '/' (ej. 'Alacant/Alicante' -> ['alacant/alicante',
    'alacant', 'alicante']) y artículos invertidos en cada parte.
    """
    raw = normalizar_texto(texto)
    if not raw:
        return []
    variantes: List[str] = []
    full = normalizar_toponimo(raw)
    if full:
        variantes.append(full)
    if "/" in raw:
        for parte in raw.split("/"):
            v = normalizar_toponimo(parte)
            if v and v not in variantes:
                variantes.append(v)
    return variantes


def quitar_articulo_inicial(toponimo_norm: str) -> str:
    """Elimina un artículo inicial conocido para comparaciones de respaldo."""
    partes = toponimo_norm.split()
    if len(partes) > 1 and partes[0] in ARTICULOS_INVERTIDOS:
        return " ".join(partes[1:])
    return toponimo_norm


def limpiar_serie_coordenadas(serie: pd.Series) -> pd.Series:
    """Convierte una serie de coordenadas con comas, espacios, puntos de miles o símbolos a float64 estándar."""
    if pd.api.types.is_bool_dtype(serie):
        return pd.Series(np.nan, index=serie.index, dtype="float64")
    if pd.api.types.is_numeric_dtype(serie):
        arr = pd.to_numeric(serie, errors="coerce").to_numpy(
            dtype="float64", na_value=np.nan
        )
        return pd.Series(arr, index=serie.index, dtype="float64").replace(
            [np.inf, -np.inf], np.nan
        )

    def _limpiar_str(val: Any) -> Any:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return pd.NA
        try:
            if pd.isna(val):
                return pd.NA
        except (TypeError, ValueError):
            pass
        s = (
            str(val)
            .replace("\u00a0", "")
            .replace("°", "")
            .replace("º", "")
            .strip()
        )
        if not s:
            return pd.NA
        if "," in s:
            if "." in s:
                s = s.replace(".", "")
            s = s.replace(",", ".")
        elif re.match(r"^[-+]?\d+(?:\.\d+){2,}$", s):
            # Varios puntos sin coma (ej. "40.136.215" por separador de miles en Excel)
            s = s.replace(".", "")
        return s

    limpia = pd.Series(
        [_limpiar_str(v) for v in serie],
        index=serie.index,
        dtype="string",
    )
    # Extraer primer número decimal con signo opcional si viene acompañado de texto
    extraida = limpia.str.extract(r"([-+]?\d+(?:\.\d+)?)", expand=False)
    arr = pd.to_numeric(extraida, errors="coerce").to_numpy(
        dtype="float64", na_value=np.nan
    )
    return pd.Series(arr, index=serie.index, dtype="float64").replace(
        [np.inf, -np.inf], np.nan
    )


def reparar_coordenadas_sin_punto(
    lat: float,
    lon: float,
    centro_ref: Optional[Tuple[float, float]] = None,
) -> Tuple[float, float, bool]:
    """
    Repara coordenadas españolas a las que Excel eliminó el punto decimal
    al interpretarlo como separador de miles (ej. 40136215 -> 40.136215).
    Devuelve (lat_reparada, lon_reparada, fue_reparada).
    """
    if pd.isna(lat) or pd.isna(lon):
        return (
            np.nan if pd.isna(lat) else float(lat),
            np.nan if pd.isna(lon) else float(lon),
            False,
        )

    lat_f = float(lat)
    lon_f = float(lon)

    if not np.isfinite(lat_f) or not np.isfinite(lon_f):
        return lat_f, lon_f, False

    if esta_en_rango_espana(lat_f, lon_f) or esta_en_rango_espana(lon_f, lat_f):
        return lat_f, lon_f, False

    if abs(lat_f) <= 90 and abs(lon_f) <= 180:
        return lat_f, lon_f, False

    # Candidatos para latitud positiva en España (2 dígitos enteros: 27.3..44.2)
    cands_lat: List[float] = []
    if 27.3 <= lat_f <= 44.2:
        cands_lat.append(lat_f)
    elif lat_f > 90:
        dig_lat = str(int(round(abs(lat_f))))
        if len(dig_lat) >= 2:
            val_lat = float(
                f"{dig_lat[:2]}.{dig_lat[2:]}" if len(dig_lat) > 2 else dig_lat
            )
            if 27.3 <= val_lat <= 44.2:
                cands_lat.append(val_lat)

    # Candidatos para longitud en España (-18.5..5.0): k in {1, 2, 0}
    cands_lon: List[Tuple[float, Optional[int]]] = []
    if -18.5 <= lon_f <= 5.0:
        cands_lon.append((lon_f, None))
    elif abs(lon_f) > 180:
        sign = -1.0 if lon_f < 0 else 1.0
        dig_lon = str(int(round(abs(lon_f))))
        for k in (1, 2, 0):
            if k == 0:
                val_lon = sign * float(f"0.{dig_lon}")
            elif len(dig_lon) >= k:
                parte_ent = dig_lon[:k]
                parte_dec = dig_lon[k:]
                val_lon = sign * float(
                    f"{parte_ent}.{parte_dec}" if parte_dec else parte_ent
                )
            else:
                continue
            if -18.5 <= val_lon <= 5.0:
                cands_lon.append((val_lon, k))

    candidatos_validos: List[Tuple[float, float, Optional[int]]] = [
        (c_lat, c_lon, k_lon)
        for c_lat in cands_lat
        for c_lon, k_lon in cands_lon
        if esta_en_rango_espana(c_lat, c_lon)
    ]

    if not candidatos_validos:
        return lat_f, lon_f, False

    if len(candidatos_validos) == 1:
        c_lat, c_lon, _ = candidatos_validos[0]
        return c_lat, c_lon, True

    if centro_ref is not None:
        ref_lat, ref_lon = centro_ref
        mejor = min(
            candidatos_validos,
            key=lambda item: (item[0] - ref_lat) ** 2 + (item[1] - ref_lon) ** 2,
        )
        return mejor[0], mejor[1], True

    def _prioridad_sin_ref(item: Tuple[float, float, Optional[int]]) -> Tuple[int, float]:
        c_lat, _, k_lon = item
        if c_lat >= 35.0:
            prio = 0 if k_lon in (1, None) else (1 if k_lon == 0 else 2)
        elif c_lat < 32.0:
            prio = 0 if k_lon in (2, None) else 1
        else:
            prio = 0
        return (prio, abs(item[1]))

    mejor = min(candidatos_validos, key=_prioridad_sin_ref)
    return mejor[0], mejor[1], True


def esta_en_rango_espana(lat: float, lon: float) -> bool:
    """
    Comprueba si (lat, lon) está dentro de los límites geográficos de España,
    incluyendo Península, Baleares, Ceuta, Melilla e Islas Canarias.
    """
    if pd.isna(lat) or pd.isna(lon):
        return False
    peninsula_baleares_ceuta_melilla = (35.0 <= lat <= 44.2) and (-10.0 <= lon <= 5.0)
    canarias = (27.3 <= lat <= 29.6) and (-18.5 <= lon <= -13.2)
    return peninsula_baleares_ceuta_melilla or canarias


def detectar_columna_coordenadas_combinadas(df: pd.DataFrame) -> Optional[str]:
    """
    Detecta si existe una única columna con pares 'lat, lon' cuando el archivo
    no separa Latitud y Longitud en dos columnas.
    """
    patron = re.compile(
        r"^\s*([-+]?\d{1,2}[.,]\d+)\s*[,;/\s]\s*([-+]?\d{1,3}[.,]\d+)\s*$"
    )
    for col in df.columns:
        muestra = df[col].dropna().astype(str).head(50)
        if muestra.empty:
            continue
        aciertos = muestra.apply(lambda x: bool(patron.match(x))).mean()
        if aciertos >= 0.6:
            return str(col)
    return None


def separar_coordenadas_combinadas(
    df: pd.DataFrame, col_combinada: str
) -> Tuple[pd.DataFrame, str, str]:
    """Separa una columna combinada 'lat, lon' en columnas Latitud y Longitud."""
    df_out = df.copy()
    patron = re.compile(
        r"([-+]?\d{1,2}[.,]\d+)\s*[,;/\s]\s*([-+]?\d{1,3}[.,]\d+)"
    )
    extraido = df_out[col_combinada].astype(str).str.extract(patron)
    col_lat = "_LATITUD_EXTRAIDA"
    col_lon = "_LONGITUD_EXTRAIDA"
    df_out[col_lat] = pd.Series(
        pd.to_numeric(
            extraido[0].str.replace(",", ".", regex=False), errors="coerce"
        ).to_numpy(dtype="float64", na_value=np.nan),
        index=df_out.index,
        dtype="float64",
    )
    df_out[col_lon] = pd.Series(
        pd.to_numeric(
            extraido[1].str.replace(",", ".", regex=False), errors="coerce"
        ).to_numpy(dtype="float64", na_value=np.nan),
        index=df_out.index,
        dtype="float64",
    )
    return df_out, col_lat, col_lon


# ==============================================================================
# AUTODETECCIÓN INTELIGENTE DE COLUMNAS
# ==============================================================================


def _puntuar_columna(col_norm: str, sinonimos: Sequence[str]) -> float:
    """Calcula una puntuación de afinidad [0..1] entre una columna y sinónimos."""
    if not col_norm:
        return 0.0

    # Nivel 1: coincidencia exacta
    for idx, sin in enumerate(sinonimos):
        if col_norm == sin:
            return 1.0 - (idx * 0.002)

    # Nivel 2: coincidencia de palabra completa o subcadena fuerte (evitando falsos positivos en sinónimos cortos)
    tokens_col = set(col_norm.split())
    mejor = 0.0
    for sin in sinonimos:
        tokens_sin = set(sin.split())
        if tokens_sin.issubset(tokens_col):
            mejor = max(mejor, 0.85)
        elif len(sin) >= 5 and sin in col_norm:
            mejor = max(mejor, 0.78)
        elif len(col_norm) >= 5 and col_norm in sin:
            mejor = max(mejor, 0.72)
        else:
            ratio = difflib.SequenceMatcher(None, col_norm, sin).ratio()
            if ratio >= 0.78:
                mejor = max(mejor, ratio * 0.85)
    return mejor


def detectar_columnas_inteligente(
    df: pd.DataFrame,
    municipios_ign_norm: Optional[set] = None,
) -> Dict[str, Optional[str]]:
    """
    Detecta de forma flexible las columnas principales y auxiliares de un DataFrame:
      - Requeridas: municipio, latitud, longitud, recurso
      - Opcionales: id, descripcion, direccion, url
    Combina coincidencia exacta normalizada, similitud fuzzy/subcadena y
    heurística de contenido (rangos de coordenadas y coincidencia de municipios).
    """
    columnas = list(df.columns)
    cols_norm = {col: normalizar_nombre_columna(col) for col in columnas}

    roles_principales = ["latitud", "longitud", "municipio", "recurso", "id"]
    roles_opcionales = ["descripcion", "direccion", "url"]

    resultado: Dict[str, Optional[str]] = {
        rol: None for rol in roles_principales + roles_opcionales
    }
    usadas: set = set()

    # 1. Emparejamiento por nombre (ordenado por mejor puntuación global)
    candidatos: List[Tuple[float, str, str]] = []
    for rol in roles_principales:
        sinonimos = SINONIMOS_COLUMNAS[rol]
        for col in columnas:
            score = _puntuar_columna(cols_norm[col], sinonimos)
            if score >= 0.65:
                candidatos.append((score, rol, col))

    candidatos.sort(key=lambda x: x[0], reverse=True)
    for score, rol, col in candidatos:
        if resultado[rol] is None and col not in usadas:
            resultado[rol] = col
            usadas.add(col)

    # 2. Heurística por contenido para Latitud y Longitud si siguen faltando
    if resultado["latitud"] is None or resultado["longitud"] is None:
        scores_lat: List[Tuple[float, str]] = []
        scores_lon: List[Tuple[float, str]] = []
        for col in columnas:
            if col in usadas:
                continue
            serie_num = limpiar_serie_coordenadas(df[col]).dropna()
            if len(serie_num) == 0:
                continue
            ratio_num = len(serie_num) / max(len(df[col].dropna()), 1)
            if ratio_num < 0.5:
                continue
            # Excluir columnas de enteros secuenciales tipo ID (1, 2, 3...)
            tiene_decimales = (serie_num % 1 != 0).mean() > 0.3
            if not tiene_decimales:
                continue
            en_lat = serie_num.between(27.0, 44.5).mean()
            en_lon = serie_num.between(-19.0, 5.0).mean()
            if en_lat >= 0.7:
                scores_lat.append((float(en_lat), col))
            if en_lon >= 0.7:
                scores_lon.append((float(en_lon), col))

        if resultado["latitud"] is None and scores_lat:
            scores_lat.sort(reverse=True)
            col_elegida = scores_lat[0][1]
            resultado["latitud"] = col_elegida
            usadas.add(col_elegida)

        if resultado["longitud"] is None and scores_lon:
            scores_lon.sort(reverse=True)
            for _, col_cand in scores_lon:
                if col_cand not in usadas:
                    resultado["longitud"] = col_cand
                    usadas.add(col_cand)
                    break

    # 3. Heurística por contenido para Municipio si sigue faltando
    if resultado["municipio"] is None and municipios_ign_norm:
        mejor_mun_score = 0.0
        mejor_mun_col: Optional[str] = None
        for col in columnas:
            if col in usadas:
                continue
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
            muestra = df[col].dropna().astype(str).head(150)
            if muestra.empty:
                continue
            aciertos = muestra.apply(
                lambda x: normalizar_toponimo(x) in municipios_ign_norm
            ).mean()
            if aciertos > mejor_mun_score and aciertos >= 0.35:
                mejor_mun_score = float(aciertos)
                mejor_mun_col = col
        if mejor_mun_col is not None:
            resultado["municipio"] = mejor_mun_col
            usadas.add(mejor_mun_col)

    # 4. Respaldo para Recurso: primera columna de texto libre no usada
    if resultado["recurso"] is None:
        for col in columnas:
            if col not in usadas and not pd.api.types.is_numeric_dtype(df[col]):
                resultado["recurso"] = col
                usadas.add(col)
                break

    # 5. Columnas opcionales para la exportación de discrepancias
    for rol in roles_opcionales:
        mejor_score = 0.0
        mejor_col: Optional[str] = None
        for col in columnas:
            if col in {
                resultado["latitud"],
                resultado["longitud"],
                resultado["municipio"],
                resultado["id"],
            }:
                continue
            score = _puntuar_columna(cols_norm[col], SINONIMOS_COLUMNAS[rol])
            if score > mejor_score and score >= 0.70:
                mejor_score = score
                mejor_col = col
        resultado[rol] = mejor_col

    return resultado


# ==============================================================================
# CARGA Y PREPARACIÓN DE CARTOGRAFÍA OFICIAL DEL IGN
# ==============================================================================


def preparar_gdf_ign(gdf_raw: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """
    Estandariza columnas de un GeoDataFrame de municipios del IGN en EPSG:4326
    con columnas: MUNICIPIO_OFICIAL_IGN, PROVINCIA_OFICIAL_IGN, COD_PROVINCIA_IGN, geometry.
    Filtra municipios históricos dados de baja si existe la columna 'fecha_baja'.
    """
    gdf = gdf_raw.copy()
    if "fecha_baja" in gdf.columns:
        mask_vigentes = gdf["fecha_baja"].isna() | (
            gdf["fecha_baja"].astype(str).str.strip() == ""
        )
        gdf = gdf[mask_vigentes].copy()

    cols_ign = list(gdf.columns)

    cand_nombre = [
        "rotulo",
        "name",
        "nombre",
        "municipio",
        "label",
        "rotulo_muni",
        "NAME",
        "MUNICIPIO_OFICIAL_IGN",
    ]
    col_ign_nombre = next((c for c in cand_nombre if c in cols_ign), None)
    if not col_ign_nombre:
        col_ign_nombre = [c for c in cols_ign if c != "geometry"][0]

    cand_cpro = [
        "id_prov",
        "cpro",
        "CPRO",
        "cod_prov",
        "codprov",
        "cpro_code",
        "ine.prov.cd",
        "COD_PROVINCIA_IGN",
    ]
    col_ign_cpro = next((c for c in cand_cpro if c in cols_ign), None)

    cand_prov = [
        "provincia",
        "PROVINCIA",
        "ine.prov.name",
        "ine.prov.name.x",
        "cpro_name",
        "prov_name",
        "PROVINCIA_OFICIAL_IGN",
    ]
    col_ign_prov = next((c for c in cand_prov if c in cols_ign), None)

    gdf["MUNICIPIO_OFICIAL_IGN"] = gdf[col_ign_nombre].astype(str)

    if col_ign_cpro:
        gdf["COD_PROVINCIA_IGN"] = (
            gdf[col_ign_cpro]
            .astype(str)
            .str.extract(r"(\d+)", expand=False)
            .fillna("00")
            .str.zfill(2)
            .str[:2]
        )
    else:
        col_ine = next(
            (
                c
                for c in ["id_ine", "inecode", "cod_ine", "codigo_ine", "id", "LAU_CODE"]
                if c in cols_ign
            ),
            None,
        )
        if col_ine:
            gdf["COD_PROVINCIA_IGN"] = (
                gdf[col_ine]
                .astype(str)
                .str.extract(r"(\d+)", expand=False)
                .fillna("00000")
                .str.zfill(5)
                .str[:2]
            )
        else:
            gdf["COD_PROVINCIA_IGN"] = "00"

    if col_ign_prov:
        gdf["PROVINCIA_OFICIAL_IGN"] = gdf[col_ign_prov].astype(str)
    else:
        gdf["PROVINCIA_OFICIAL_IGN"] = (
            gdf["COD_PROVINCIA_IGN"].map(PROVINCIAS_INE).fillna("Desconocida")
        )

    cols_keep = [
        "MUNICIPIO_OFICIAL_IGN",
        "PROVINCIA_OFICIAL_IGN",
        "COD_PROVINCIA_IGN",
        "geometry",
    ]
    if gdf.crs is None:
        gdf = gdf.set_crs(epsg=4326)
    gdf_4326 = gdf[cols_keep].to_crs(epsg=4326)
    return gdf_4326


@st.cache_resource(show_spinner=False)
def cargar_cartografia_ign(url: str = URL_IGN_MUNICIPIOS) -> gpd.GeoDataFrame:
    """
    Descarga y cachea en memoria la cartografía municipal oficial del IGN,
    incluyendo opcionalmente la capa oficial de Canarias si no viene en el archivo principal.
    """
    gdf_raw = gpd.read_file(url, engine="pyogrio")
    gdf_pen = preparar_gdf_ign(gdf_raw)

    provs_presentes = set(gdf_pen["COD_PROVINCIA_IGN"].dropna().unique())
    if url == URL_IGN_MUNICIPIOS and not any(
        cp in provs_presentes for cp in PROVINCIAS_CANARIAS
    ):
        try:
            gdf_raw_can = gpd.read_file(URL_IGN_MUNICIPIOS_CANARIAS, engine="pyogrio")
            if not gdf_raw_can.empty:
                gdf_can = preparar_gdf_ign(gdf_raw_can)
                gdf_combinado = pd.concat([gdf_pen, gdf_can], ignore_index=True)
                return gpd.GeoDataFrame(
                    gdf_combinado, geometry="geometry", crs="EPSG:4326"
                )
        except Exception:
            pass

    return gdf_pen


def obtener_set_municipios_ign(gdf_espana_4326: gpd.GeoDataFrame) -> set:
    """Construye un conjunto de variantes normalizadas de todos los municipios del IGN."""
    res: set = set()
    for nom in gdf_espana_4326["MUNICIPIO_OFICIAL_IGN"].dropna().unique():
        for v in variantes_toponimo(nom):
            res.add(v)
    return res


# ==============================================================================
# MOTOR DE AUDITORÍA GEOESPACIAL
# ==============================================================================


def _distancia_metrica_4326(punto_4326: Any, poligono_4326: Any, cod_prov: str) -> float:
    """
    Calcula la distancia en metros entre un punto y un polígono en WGS84,
    proyectando dinámicamente a UTM 28N para Canarias y UTM 30N para el resto.
    """
    usar_canarias = (cod_prov in PROVINCIAS_CANARIAS) or (punto_4326.y < 32.0)
    trans = _TRANS_4326_TO_25828 if usar_canarias else _TRANS_4326_TO_25830

    def _transformar(coords: np.ndarray) -> np.ndarray:
        if len(coords) == 0:
            return coords
        return np.column_stack(trans.transform(coords[:, 0], coords[:, 1]))

    pt_utm = shapely.transform(punto_4326, _transformar)
    poly_utm = shapely.transform(poligono_4326, _transformar)
    return round(float(pt_utm.distance(poly_utm)), 1)


def _construir_indice_poligonos(
    gdf_4326: gpd.GeoDataFrame,
) -> Tuple[
    Dict[str, Dict[str, List[Tuple[str, Any]]]],
    Dict[str, List[Tuple[str, str, Any]]],
]:
    """
    Construye índices rápidos de búsqueda de polígonos:
      - por_provincia[cod_prov][variante_norm] -> [(nombre_oficial, geom), ...]
      - nacional[variante_norm] -> [(cod_prov, nombre_oficial, geom), ...]
    """
    por_provincia: Dict[str, Dict[str, List[Tuple[str, Any]]]] = {}
    nacional: Dict[str, List[Tuple[str, str, Any]]] = {}

    for nom, cpro, geom in zip(
        gdf_4326["MUNICIPIO_OFICIAL_IGN"],
        gdf_4326["COD_PROVINCIA_IGN"],
        gdf_4326["geometry"],
    ):
        cpro_str = str(cpro).zfill(2)
        dicc_prov = por_provincia.setdefault(cpro_str, {})
        for var in variantes_toponimo(nom):
            dicc_prov.setdefault(var, []).append((nom, geom))
            nacional.setdefault(var, []).append((cpro_str, nom, geom))
    return por_provincia, nacional


def _buscar_poligono_declarado(
    mun_excel_raw: Any,
    punto_4326: Any,
    codigo_provincia_auditada: Optional[str],
    por_provincia: Dict[str, Dict[str, List[Tuple[str, Any]]]],
    nacional: Dict[str, List[Tuple[str, str, Any]]],
) -> Optional[Tuple[str, Any]]:
    """
    Localiza el polígono del municipio declarado priorizando la provincia auditada
    para resolver homónimos correctamente (ej. Aldea del Obispo en Salamanca vs
    La Aldea del Obispo en Cáceres).
    Devuelve (cod_prov, geometry) o None si no se localiza.
    """
    variantes_exc = variantes_toponimo(mun_excel_raw)
    if not variantes_exc:
        return None
    sin_art_exc = [quitar_articulo_inicial(v) for v in variantes_exc]

    def _elegir_mas_cercano(candidatos: List[Tuple[str, Any]]) -> Tuple[str, Any]:
        if len(candidatos) == 1:
            return candidatos[0]
        return min(
            candidatos,
            key=lambda item: punto_4326.distance(item[1]),
        )

    # 1. Prioridad absoluta: provincia auditada (si está definida)
    if codigo_provincia_auditada and codigo_provincia_auditada in por_provincia:
        dicc_prov = por_provincia[codigo_provincia_auditada]

        # 1a. Coincidencia exacta de variante
        for v in variantes_exc:
            if v in dicc_prov:
                cands = [(codigo_provincia_auditada, g) for _, g in dicc_prov[v]]
                return _elegir_mas_cercano(cands)

        # 1b. Coincidencia ignorando artículo inicial dentro de la misma provincia
        cands_art: List[Tuple[str, Any]] = []
        for k, lista in dicc_prov.items():
            if quitar_articulo_inicial(k) in sin_art_exc:
                cands_art.extend((codigo_provincia_auditada, g) for _, g in lista)
        if cands_art:
            return _elegir_mas_cercano(cands_art)

        # 1c. Coincidencia parcial segura (mínimo 4 caracteres) dentro de la provincia
        cands_parc: List[Tuple[str, Any]] = []
        for v in variantes_exc:
            if len(v) < 4:
                continue
            for k, lista in dicc_prov.items():
                if len(k) >= 4 and (v in k or k in v):
                    cands_parc.extend((codigo_provincia_auditada, g) for _, g in lista)
        if cands_parc:
            return _elegir_mas_cercano(cands_parc)

    # 2. Búsqueda nacional (si no estaba en la provincia o en modo 'Toda España')
    cands_nac: List[Tuple[str, Any]] = []
    for v in variantes_exc:
        if v in nacional:
            cands_nac.extend((cp, g) for cp, _, g in nacional[v])
    if cands_nac:
        return _elegir_mas_cercano(cands_nac)

    # 2b. Nacional ignorando artículo inicial
    cands_nac_art: List[Tuple[str, Any]] = []
    for k, lista in nacional.items():
        if quitar_articulo_inicial(k) in sin_art_exc:
            cands_nac_art.extend((cp, g) for cp, _, g in lista)
    if cands_nac_art:
        return _elegir_mas_cercano(cands_nac_art)

    # 2c. Nacional parcial segura (mínimo 5 caracteres para evitar falsos cruces)
    cands_nac_parc: List[Tuple[str, Any]] = []
    for v in variantes_exc:
        if len(v) < 5:
            continue
        for k, lista in nacional.items():
            if len(k) >= 5 and (v in k or k in v):
                cands_nac_parc.extend((cp, g) for cp, _, g in lista)
    if cands_nac_parc:
        return _elegir_mas_cercano(cands_nac_parc)

    return None


def detectar_provincia_automatica(
    df: pd.DataFrame,
    col_lat: str,
    col_lon: str,
    col_municipio: str,
    gdf_espana_4326: gpd.GeoDataFrame,
) -> Tuple[Optional[str], str, int, int]:
    """
    Detecta la provincia más probable del catálogo usando los puntos GPS válidos
    (o el listado de municipios declarados si no hubiera coordenadas válidas).
    Devuelve (codigo_ine, nombre_provincia, puntos_en_provincia, total_evaluados).
    """
    lats = limpiar_serie_coordenadas(df[col_lat])
    lons = limpiar_serie_coordenadas(df[col_lon])
    mask = lats.notna() & lons.notna()

    if mask.any():
        gdf_pts = gpd.GeoDataFrame(
            {"_idx": df.index[mask]},
            geometry=gpd.points_from_xy(lons[mask], lats[mask]),
            crs="EPSG:4326",
        )
        joined = gpd.sjoin(
            gdf_pts,
            gdf_espana_4326[["COD_PROVINCIA_IGN", "geometry"]],
            how="inner",
            predicate="within",
        )
        joined = joined.drop_duplicates(subset=["_idx"])
        if not joined.empty:
            conteo = joined["COD_PROVINCIA_IGN"].value_counts()
            cod = str(conteo.index[0]).zfill(2)
            nombre = PROVINCIAS_INE.get(cod, cod)
            return cod, nombre, int(conteo.iloc[0]), int(len(joined))

    # Respaldo por municipios declarados si las coordenadas no cruzaron
    if col_municipio in df.columns:
        _, nacional = _construir_indice_poligonos(gdf_espana_4326)
        votos: Dict[str, int] = {}
        total = 0
        for val in df[col_municipio].dropna():
            vars_val = variantes_toponimo(val)
            encontrado = False
            for v in vars_val:
                if v in nacional:
                    for cp, _, _ in nacional[v]:
                        votos[cp] = votos.get(cp, 0) + 1
                    encontrado = True
                    break
            if encontrado:
                total += 1
        if votos:
            cod = max(votos, key=votos.get)
            return cod, PROVINCIAS_INE.get(cod, cod), votos[cod], total

    return None, "Toda España", 0, 0


def ejecutar_auditoria(
    df_origen: pd.DataFrame,
    gdf_espana_4326: gpd.GeoDataFrame,
    col_municipio: str,
    col_latitud: str,
    col_longitud: str,
    col_recurso: Optional[str] = None,
    col_id: Optional[str] = None,
    codigo_provincia_ine: Optional[str] = None,
    umbrales: Optional[Dict[str, float]] = None,
    reparar_sin_punto: bool = True,
) -> pd.DataFrame:
    """
    Ejecuta la auditoría geoespacial completa sobre el DataFrame de entrada.

    - codigo_provincia_ine: código de 2 dígitos (ej. '10') o None/'00' para Toda España.
    - umbrales: dict con claves 'fronterizo', 'borde', 'moderado', 'notable' (en metros).
    - reparar_sin_punto: si True, repara coordenadas españolas sin punto decimal (ej. 40136215 -> 40.136215).
    """
    if umbrales is None:
        umbrales = {
            "fronterizo": 100.0,
            "borde": 500.0,
            "moderado": 2000.0,
            "notable": 6000.0,
        }

    u_front = float(umbrales["fronterizo"])
    u_borde = float(umbrales["borde"])
    u_mod = float(umbrales["moderado"])
    u_not = float(umbrales["notable"])

    cod_prov_objetivo = (
        str(codigo_provincia_ine).zfill(2)
        if codigo_provincia_ine and str(codigo_provincia_ine) not in ("00", "AUTO")
        else None
    )

    df = df_origen.copy().reset_index(drop=True)

    # Limpieza segura de coordenadas en columnas de trabajo y originales
    df[col_latitud] = limpiar_serie_coordenadas(df[col_latitud])
    df[col_longitud] = limpiar_serie_coordenadas(df[col_longitud])
    df["COORD_REPARADA"] = pd.Series(False, index=df.index, dtype="bool")

    if reparar_sin_punto:
        centro_ref: Optional[Tuple[float, float]] = None
        if cod_prov_objetivo:
            gdf_prov = gdf_espana_4326[
                gdf_espana_4326["COD_PROVINCIA_IGN"] == cod_prov_objetivo
            ]
            if not gdf_prov.empty:
                bounds = gdf_prov.total_bounds
                centro_ref = (
                    float((bounds[1] + bounds[3]) / 2.0),
                    float((bounds[0] + bounds[2]) / 2.0),
                )
        if centro_ref is None:
            mask_esp_valida = [
                esta_en_rango_espana(la, lo)
                for la, lo in zip(df[col_latitud], df[col_longitud])
            ]
            if any(mask_esp_valida):
                centro_ref = (
                    float(df.loc[mask_esp_valida, col_latitud].median()),
                    float(df.loc[mask_esp_valida, col_longitud].median()),
                )

        lats_nuevas: List[float] = []
        lons_nuevas: List[float] = []
        reparadas_flags: List[bool] = []
        for la, lo in zip(df[col_latitud], df[col_longitud]):
            n_la, n_lo, rep = reparar_coordenadas_sin_punto(
                la, lo, centro_ref=centro_ref
            )
            lats_nuevas.append(np.nan if pd.isna(n_la) else float(n_la))
            lons_nuevas.append(np.nan if pd.isna(n_lo) else float(n_lo))
            reparadas_flags.append(bool(rep))

        df[col_latitud] = pd.Series(lats_nuevas, index=df.index, dtype="float64")
        df[col_longitud] = pd.Series(lons_nuevas, index=df.index, dtype="float64")
        df["COORD_REPARADA"] = pd.Series(reparadas_flags, index=df.index, dtype="bool")

    # Diagnóstico de inversión Lat/Lon (compatible con Península, Baleares y Canarias)
    def _evaluar_rango(lat: float, lon: float) -> str:
        if pd.isna(lat) or pd.isna(lon):
            return "Sin coordenadas"
        if esta_en_rango_espana(lat, lon):
            return "OK"
        if esta_en_rango_espana(lon, lat):
            return "Posible inversión Lat/Lon"
        return "Fuera de rango España"

    df["ALERTA_COORDENADAS"] = [
        _evaluar_rango(la, lo)
        for la, lo in zip(df[col_latitud], df[col_longitud])
    ]

    # Inicializar columnas resultado con tipos explícitos (evita FutureWarning de Pandas)
    df["MUNICIPIO_DETECTADO_IGN"] = pd.Series(pd.NA, index=df.index, dtype="object")
    df["PROVINCIA_DETECTADA_IGN"] = pd.Series(pd.NA, index=df.index, dtype="object")
    df["COD_PROVINCIA_DETECTADA_IGN"] = pd.Series(pd.NA, index=df.index, dtype="object")
    df["METROS_AL_LIMITE"] = pd.Series(np.nan, index=df.index, dtype="float64")
    df["ESTADO_AUDITORIA"] = pd.Series("SIN_COORDENADAS", index=df.index, dtype="object")
    df["ALERTA_INTERPROVINCIAL"] = pd.Series(
        "Sin coordenadas", index=df.index, dtype="object"
    )

    mask_validas = df[col_latitud].notna() & df[col_longitud].notna()
    if not mask_validas.any():
        return df

    # Spatial join en EPSG:4326 (exacto topológicamente en Península y Canarias)
    gdf_pts_4326 = gpd.GeoDataFrame(
        df.loc[mask_validas, [col_municipio]].copy(),
        geometry=gpd.points_from_xy(
            df.loc[mask_validas, col_longitud],
            df.loc[mask_validas, col_latitud],
        ),
        crs="EPSG:4326",
    )
    gdf_pts_4326["_orig_idx"] = gdf_pts_4326.index

    # Usamos predicate="intersects" para no perder puntos exactamente sobre la línea fronteriza
    gdf_joined = gpd.sjoin(
        gdf_pts_4326,
        gdf_espana_4326,
        how="left",
        predicate="intersects",
    )

    # Resolver duplicados si un punto cae exactamente en la frontera compartida de 2 municipios:
    # priorizar el polígono que coincide con el municipio declarado en el Excel.
    if gdf_joined["_orig_idx"].duplicated().any():
        def _coincide_declarado(row: pd.Series) -> int:
            v_exc = set(variantes_toponimo(row.get(col_municipio, "")))
            v_ign = set(variantes_toponimo(row.get("MUNICIPIO_OFICIAL_IGN", "")))
            return 0 if (v_exc and v_ign and v_exc.intersection(v_ign)) else 1

        gdf_joined["_prioridad_dup"] = gdf_joined.apply(_coincide_declarado, axis=1)
        gdf_joined = (
            gdf_joined.sort_values(by=["_orig_idx", "_prioridad_dup"])
            .drop_duplicates(subset=["_orig_idx"], keep="first")
            .drop(columns=["_prioridad_dup"])
        )

    por_provincia, nacional = _construir_indice_poligonos(gdf_espana_4326)

    estados: Dict[int, str] = {}
    distancias: Dict[int, float] = {}
    mun_ign_map: Dict[int, Any] = {}
    prov_ign_map: Dict[int, Any] = {}
    cpro_ign_map: Dict[int, Any] = {}
    alertas_prov: Dict[int, str] = {}

    for _, row in gdf_joined.iterrows():
        idx = int(row["_orig_idx"])
        mun_excel_raw = row.get(col_municipio, "")
        mun_ign_raw = row.get("MUNICIPIO_OFICIAL_IGN", pd.NA)
        prov_ign_raw = row.get("PROVINCIA_OFICIAL_IGN", pd.NA)
        cod_prov_ign_raw = row.get("COD_PROVINCIA_IGN", pd.NA)
        punto_4326 = row["geometry"]

        cod_prov_ign = (
            str(cod_prov_ign_raw).zfill(2) if pd.notna(cod_prov_ign_raw) else ""
        )

        # Caso 1: Fuera de cualquier polígono español del IGN
        if pd.isna(mun_ign_raw):
            estados[idx] = "FUERA_DE_PROVINCIA_O_SIN_POLIGONO"
            distancias[idx] = np.nan
            mun_ign_map[idx] = pd.NA
            prov_ign_map[idx] = pd.NA
            cpro_ign_map[idx] = pd.NA
            alertas_prov[idx] = "Sin polígono coincidente"
            continue

        # Evaluar alerta interprovincial
        if cod_prov_objetivo and cod_prov_ign and cod_prov_ign != cod_prov_objetivo:
            alerta_prov = f"⚠️ OTRA PROVINCIA ({prov_ign_raw})"
        elif not cod_prov_objetivo:
            alerta_prov = f"Provincia: {prov_ign_raw}"
        else:
            alerta_prov = "Misma provincia"

        mun_ign_map[idx] = mun_ign_raw
        prov_ign_map[idx] = prov_ign_raw
        cpro_ign_map[idx] = cod_prov_ign
        alertas_prov[idx] = alerta_prov

        # Caso 2: Coincidencia de topónimo (resolviendo artículos invertidos y cooficiales)
        vars_excel = set(variantes_toponimo(mun_excel_raw))
        vars_ign = set(variantes_toponimo(mun_ign_raw))

        coincide_nombre = bool(vars_excel and vars_ign and vars_excel.intersection(vars_ign))
        # Si hay provincia objetivo y el punto está en esa misma provincia, también aceptar
        # omisión de artículo cuando ambos tienen el mismo núcleo (ej. "Aldea del Obispo" en Cáceres)
        if not coincide_nombre and vars_excel and vars_ign:
            if (not cod_prov_objetivo) or (cod_prov_ign == cod_prov_objetivo):
                sin_art_exc = {quitar_articulo_inicial(v) for v in vars_excel}
                sin_art_ign = {quitar_articulo_inicial(v) for v in vars_ign}
                if sin_art_exc.intersection(sin_art_ign):
                    coincide_nombre = True

        if coincide_nombre:
            estados[idx] = "CORRECTO"
            distancias[idx] = 0.0
            continue

        # Caso 3: Discrepancia -> calcular distancia métrica al polígono del municipio declarado
        busqueda = _buscar_poligono_declarado(
            mun_excel_raw=mun_excel_raw,
            punto_4326=punto_4326,
            codigo_provincia_auditada=cod_prov_objetivo,
            por_provincia=por_provincia,
            nacional=nacional,
        )

        if busqueda is None:
            distancia_m = np.nan
            categoria = "TOPÓNIMO_NO_LOCALIZADO"
        else:
            cod_prov_poly, poly_4326 = busqueda
            distancia_m = _distancia_metrica_4326(
                punto_4326, poly_4326, cod_prov_poly
            )
            if distancia_m == 0.0:
                categoria = "CORRECTO"
            elif distancia_m <= u_front:
                categoria = "FRONTERIZO_INMEDIATO"
            elif distancia_m <= u_borde:
                categoria = "BORDE_CERCANO"
            elif distancia_m <= u_mod:
                categoria = "DESVÍO_MODERADO"
            elif distancia_m <= u_not:
                categoria = "DESVÍO_NOTABLE"
            else:
                categoria = "ERROR_GRAVE"

        estados[idx] = categoria
        distancias[idx] = distancia_m

    idx_list = list(estados.keys())
    if idx_list:
        df.loc[idx_list, "ESTADO_AUDITORIA"] = [estados[i] for i in idx_list]
        df.loc[idx_list, "METROS_AL_LIMITE"] = [distancias[i] for i in idx_list]
        df.loc[idx_list, "MUNICIPIO_DETECTADO_IGN"] = [mun_ign_map[i] for i in idx_list]
        df.loc[idx_list, "PROVINCIA_DETECTADA_IGN"] = [prov_ign_map[i] for i in idx_list]
        df.loc[idx_list, "COD_PROVINCIA_DETECTADA_IGN"] = [
            cpro_ign_map[i] for i in idx_list
        ]
        df.loc[idx_list, "ALERTA_INTERPROVINCIAL"] = [alertas_prov[i] for i in idx_list]

    return df


# ==============================================================================
# MAPA INTERACTIVO FOLIUM (UNIFICADO ANTI-SOLAPAMIENTO)
# ==============================================================================


def _formatear_km(metros: float) -> str:
    km = metros / 1000.0
    return f"{km:.0f} km" if km.is_integer() else f"{km:.1f} km"


def generar_mapa_folium(
    df_auditado: pd.DataFrame,
    gdf_espana_4326: gpd.GeoDataFrame,
    col_municipio: str,
    col_latitud: str,
    col_longitud: str,
    col_recurso: Optional[str] = None,
    col_id: Optional[str] = None,
    codigo_provincia_ine: Optional[str] = None,
    nombre_provincia: str = "España",
    umbrales: Optional[Dict[str, float]] = None,
    mostrar_correctos: bool = True,
    simplificar_geometria: float = 0.0005,
) -> Tuple[folium.Map, int]:
    """
    Genera el mapa interactivo Folium consolidado con:
      - Mapa base Esri World Street Map
      - Contornos municipales oficiales del IGN (simplificados ligeramente para rapidez)
      - 8 capas por categoría con recuentos dinámicos
      - Agrupación anti-solapamiento por (lat, lon) priorizando la categoría más grave
      - Popups HTML con todos los activos y alertas interprovinciales
    Devuelve (mapa_folium, puntos_excluidos_fuera_rango).
    """
    if umbrales is None:
        umbrales = {
            "fronterizo": 100.0,
            "borde": 500.0,
            "moderado": 2000.0,
            "notable": 6000.0,
        }

    u_front = float(umbrales["fronterizo"])
    u_borde = float(umbrales["borde"])
    u_mod = float(umbrales["moderado"])
    u_not = float(umbrales["notable"])

    df_mapa = df_auditado[
        df_auditado[col_latitud].notna() & df_auditado[col_longitud].notna()
    ].copy()

    if not mostrar_correctos:
        df_mapa = df_mapa[df_mapa["ESTADO_AUDITORIA"] != "CORRECTO"]

    # Filtrar puntos dentro de los rangos geográficos de España (Península, Baleares, Canarias)
    mask_en_espana = [
        esta_en_rango_espana(la, lo)
        for la, lo in zip(df_mapa[col_latitud], df_mapa[col_longitud])
    ]
    df_mapa_valido = df_mapa[mask_en_espana].copy()
    puntos_fuera_rango = int(len(df_mapa) - len(df_mapa_valido))

    # Seleccionar capa de contornos municipales a dibujar
    cod_prov = (
        str(codigo_provincia_ine).zfill(2)
        if codigo_provincia_ine and str(codigo_provincia_ine) not in ("00", "AUTO")
        else None
    )

    if cod_prov:
        gdf_contornos = gdf_espana_4326[
            gdf_espana_4326["COD_PROVINCIA_IGN"] == cod_prov
        ].copy()
        titulo_capa_ign = f"Términos Municipales {nombre_provincia} (IGN)"
    else:
        provs_presentes = (
            df_auditado["COD_PROVINCIA_DETECTADA_IGN"].dropna().unique().tolist()
        )
        if provs_presentes:
            gdf_contornos = gdf_espana_4326[
                gdf_espana_4326["COD_PROVINCIA_IGN"].isin(provs_presentes)
            ].copy()
            titulo_capa_ign = "Términos Municipales Detectados (IGN)"
        else:
            gdf_contornos = gdf_espana_4326.iloc[0:0].copy()
            titulo_capa_ign = "Términos Municipales (IGN)"

    # Centro dinámico y zoom inicial
    if not df_mapa_valido.empty:
        centro_lat = float(df_mapa_valido[col_latitud].median())
        centro_lon = float(df_mapa_valido[col_longitud].median())
        zoom_inicial = 9 if cod_prov else 6
    elif not gdf_contornos.empty:
        bounds = gdf_contornos.total_bounds  # minx, miny, maxx, maxy
        centro_lon = float((bounds[0] + bounds[2]) / 2.0)
        centro_lat = float((bounds[1] + bounds[3]) / 2.0)
        zoom_inicial = 9 if cod_prov else 6
    else:
        centro_lat, centro_lon, zoom_inicial = 40.4168, -3.7038, 6

    mapa = folium.Map(
        location=[centro_lat, centro_lon],
        zoom_start=zoom_inicial,
        tiles=None,
        width="100%",
        height="100%",
    )
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}",
        attr="Esri & IGN",
        name="Esri World Street Map",
        control=False,
    ).add_to(mapa)

    # Añadir contornos municipales (simplificación ligera solo para renderizado rápido)
    if not gdf_contornos.empty:
        if simplificar_geometria and simplificar_geometria > 0:
            gdf_contornos["geometry"] = gdf_contornos.geometry.simplify(
                simplificar_geometria, preserve_topology=True
            )
        folium.GeoJson(
            gdf_contornos[["MUNICIPIO_OFICIAL_IGN", "PROVINCIA_OFICIAL_IGN", "geometry"]],
            name=titulo_capa_ign,
            style_function=lambda _x: {
                "fillColor": "#2c3e50",
                "color": "#34495e",
                "weight": 1.0,
                "fillOpacity": 0.04,
            },
            tooltip=folium.GeoJsonTooltip(
                fields=["MUNICIPIO_OFICIAL_IGN", "PROVINCIA_OFICIAL_IGN"],
                aliases=["Municipio IGN:", "Provincia:"],
            ),
        ).add_to(mapa)

    conteos = df_auditado["ESTADO_AUDITORIA"].value_counts().to_dict()

    capas: Dict[str, folium.FeatureGroup] = {
        "ERROR_GRAVE": folium.FeatureGroup(
            name=f"🔴 Error Grave (> {_formatear_km(u_not)}) [{conteos.get('ERROR_GRAVE', 0)}]",
            show=True,
        ),
        "FUERA_DE_PROVINCIA_O_SIN_POLIGONO": folium.FeatureGroup(
            name=f"🟣 Fuera de Provincia / Sin Polígono [{conteos.get('FUERA_DE_PROVINCIA_O_SIN_POLIGONO', 0)}]",
            show=True,
        ),
        "DESVÍO_NOTABLE": folium.FeatureGroup(
            name=f"🟤 Desvío Notable ({_formatear_km(u_mod)} - {_formatear_km(u_not)}) [{conteos.get('DESVÍO_NOTABLE', 0)}]",
            show=True,
        ),
        "DESVÍO_MODERADO": folium.FeatureGroup(
            name=f"🟠 Desvío Moderado ({int(u_borde)}m - {_formatear_km(u_mod)}) [{conteos.get('DESVÍO_MODERADO', 0)}]",
            show=True,
        ),
        "BORDE_CERCANO": folium.FeatureGroup(
            name=f"🟡 Borde Cercano ({int(u_front)}m - {int(u_borde)}m) [{conteos.get('BORDE_CERCANO', 0)}]",
            show=True,
        ),
        "FRONTERIZO_INMEDIATO": folium.FeatureGroup(
            name=f"🟢 Fronterizo Inmediato (<= {int(u_front)}m) [{conteos.get('FRONTERIZO_INMEDIATO', 0)}]",
            show=False,
        ),
        "TOPÓNIMO_NO_LOCALIZADO": folium.FeatureGroup(
            name=f"⚪ Topónimo No Localizado [{conteos.get('TOPÓNIMO_NO_LOCALIZADO', 0)}]",
            show=False,
        ),
        "CORRECTO": folium.FeatureGroup(
            name=f"🔵 Correcto (Sin Discrepancias) [{conteos.get('CORRECTO', 0)}]",
            show=mostrar_correctos,
        ),
    }

    # Agrupación anti-solapamiento por coordenada GPS exacta
    if not df_mapa_valido.empty:
        grupos = df_mapa_valido.groupby([col_latitud, col_longitud], sort=False)

        for (lat, lon), grupo in grupos:
            estados_grupo = set(grupo["ESTADO_AUDITORIA"].values)
            estado_principal = next(
                (est for est in PRIORIDAD_ESTADOS if est in estados_grupo),
                "CORRECTO",
            )
            cfg = ESTILOS_MAPA.get(
                estado_principal, ESTILOS_MAPA["TOPÓNIMO_NO_LOCALIZADO"]
            )
            capa_destino = capas.get(estado_principal, capas["CORRECTO"])

            num_activos = len(grupo)
            primer_recurso = (
                str(grupo.iloc[0][col_recurso])
                if col_recurso
                and col_recurso in grupo.columns
                and pd.notna(grupo.iloc[0][col_recurso])
                else "Activo"
            )
            encabezado = (
                f"<b>{num_activos} Activo(s) en este punto GPS</b>"
                if num_activos > 1
                else f"<b>{html.escape(primer_recurso)}</b>"
            )

            html_items: List[str] = []
            for fila_idx, item in grupo.iterrows():
                if col_id and col_id in item.index and pd.notna(item[col_id]):
                    id_txt = f"[{html.escape(str(item[col_id]))}] "
                else:
                    try:
                        fila_label = f"Fila {int(fila_idx) + 2}"
                    except (TypeError, ValueError):
                        fila_label = f"Fila {html.escape(str(fila_idx))}"
                    id_txt = f"[{fila_label}] "

                rec_val = (
                    html.escape(str(item[col_recurso]))
                    if col_recurso
                    and col_recurso in item.index
                    and pd.notna(item[col_recurso])
                    else "Sin nombre"
                )
                mun_exc = (
                    html.escape(str(item[col_municipio]))
                    if pd.notna(item.get(col_municipio))
                    else "N/A"
                )
                mun_ign = (
                    html.escape(str(item["MUNICIPIO_DETECTADO_IGN"]))
                    if pd.notna(item.get("MUNICIPIO_DETECTADO_IGN"))
                    else "Desconocido"
                )
                prov_ign = (
                    html.escape(str(item["PROVINCIA_DETECTADA_IGN"]))
                    if pd.notna(item.get("PROVINCIA_DETECTADA_IGN"))
                    else ""
                )
                dist_m = item.get("METROS_AL_LIMITE", np.nan)
                alerta_prov = str(item.get("ALERTA_INTERPROVINCIAL", ""))
                estado_item = html.escape(str(item.get("ESTADO_AUDITORIA", "")))

                if pd.notna(dist_m) and dist_m > 0:
                    dist_txt = f"{dist_m:.1f} m"
                elif pd.notna(dist_m) and dist_m == 0:
                    dist_txt = "0 m"
                else:
                    dist_txt = "N/D"

                prov_tag = (
                    f" <b style='color:#c0392b;'>[{prov_ign}]</b>"
                    if "OTRA PROVINCIA" in alerta_prov
                    else ""
                )

                html_items.append(
                    f"""
                    <div style="border-top:1px solid #ecf0f1; padding:5px 0; margin-top:4px;">
                        <b>• {id_txt}{rec_val}</b><br />
                        <i>Excel:</i> {mun_exc} &rarr; <i>IGN:</i> <b>{mun_ign}</b>{prov_tag}<br />
                        <small style="color:#566573;">Desvío: {dist_txt} | Estado: {estado_item}</small>
                    </div>
                    """
                )

            popup_html = f"""
            <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
                        font-size: 12px; max-height: 260px; overflow-y: auto; line-height: 1.35;">
                <div style="font-size: 13px; color: #2c3e50; margin-bottom: 4px;">
                    {encabezado}
                </div>
                {"".join(html_items)}
                <div style="border-top: 1px solid #d5dbdb; margin-top: 6px; padding-top: 4px;
                            color: #7f8c8d; font-size: 11px;">
                    GPS: {lat:.5f}, {lon:.5f}
                </div>
            </div>
            """

            tooltip_txt = (
                f"[{num_activos} recurso(s)] {primer_recurso} | Estado: {estado_principal}"
            )

            folium.CircleMarker(
                location=[float(lat), float(lon)],
                radius=cfg["r"] + (2 if num_activos > 1 else 0),
                color=cfg["color"],
                fill=True,
                fill_color=cfg["fill"],
                fill_opacity=0.85,
                weight=1.5,
                popup=folium.Popup(popup_html, max_width=320),
                tooltip=html.escape(tooltip_txt),
            ).add_to(capa_destino)

    for capa in capas.values():
        capa.add_to(mapa)

    folium.LayerControl(position="topright", collapsed=False).add_to(mapa)

    return mapa, puntos_fuera_rango


# ==============================================================================
# CONSTRUCCIÓN DE INFORMES Y EXPORTACIÓN EXCEL
# ==============================================================================


def construir_tabla_coordenadas_repetidas(
    df_auditado: pd.DataFrame,
    col_municipio: str,
    col_latitud: str,
    col_longitud: str,
    col_recurso: Optional[str] = None,
    col_id: Optional[str] = None,
    solo_incidencias: bool = True,
) -> pd.DataFrame:
    """Genera la tabla de inspección rápida de coordenadas repetidas (Celda 13)."""
    df_base = df_auditado.copy()
    if solo_incidencias:
        df_base = df_base[df_base["ESTADO_AUDITORIA"] != "CORRECTO"]

    df_base = df_base[df_base[col_latitud].notna() & df_base[col_longitud].notna()].copy()
    if df_base.empty:
        return pd.DataFrame()

    df_base["VECES_REPETIDA_COORD"] = df_base.groupby(
        [col_latitud, col_longitud]
    )[col_latitud].transform("size")

    cols = []
    if col_id and col_id in df_base.columns:
        cols.append(col_id)
    if col_recurso and col_recurso in df_base.columns:
        cols.append(col_recurso)
    cols.extend(
        [
            col_municipio,
            "MUNICIPIO_DETECTADO_IGN",
            "PROVINCIA_DETECTADA_IGN",
            "ESTADO_AUDITORIA",
            "METROS_AL_LIMITE",
            col_latitud,
            col_longitud,
            "VECES_REPETIDA_COORD",
        ]
    )
    cols = [c for c in cols if c in df_base.columns]
    return (
        df_base[cols]
        .sort_values(
            by=["VECES_REPETIDA_COORD", col_latitud, col_longitud],
            ascending=[False, True, True],
        )
        .reset_index(drop=True)
    )


def construir_dataframe_exportacion(
    df_filtrado: pd.DataFrame,
    col_municipio: str,
    col_latitud: str,
    col_longitud: str,
    col_recurso: Optional[str] = None,
    col_id: Optional[str] = None,
    cols_extra: Optional[Dict[str, Optional[str]]] = None,
    incluir_todas_columnas: bool = False,
) -> pd.DataFrame:
    """
    Construye el DataFrame estructurado de discrepancias para exportar a Excel,
    replicando el formato del notebook y añadiendo columnas de diagnóstico IGN.
    """
    cols_extra = cols_extra or {}
    cols_disp = set(df_filtrado.columns)

    if incluir_todas_columnas:
        return df_filtrado.copy().reset_index(drop=True)

    df_exp = pd.DataFrame(index=df_filtrado.index)

    if col_id and col_id in cols_disp:
        df_exp["ID"] = df_filtrado[col_id]
    else:
        df_exp["ID"] = df_filtrado.index + 2  # Nº fila aproximado en Excel

    df_exp["NOMBRE ACTIVO"] = (
        df_filtrado[col_recurso]
        if col_recurso and col_recurso in cols_disp
        else ""
    )

    col_desc = cols_extra.get("descripcion")
    if col_desc and col_desc in cols_disp:
        df_exp["DESCRIPCION_CORTA"] = df_filtrado[col_desc]

    col_dir = cols_extra.get("direccion")
    if col_dir and col_dir in cols_disp:
        df_exp["DIRECCIÓN"] = df_filtrado[col_dir]

    df_exp["LATITUD"] = df_filtrado[col_latitud] if col_latitud in cols_disp else ""
    df_exp["LONGITUD"] = df_filtrado[col_longitud] if col_longitud in cols_disp else ""

    col_url = cols_extra.get("url")
    if col_url and col_url in cols_disp:
        df_exp["URL FUENTE"] = df_filtrado[col_url]

    df_exp["Municipio Declarado"] = (
        df_filtrado[col_municipio] if col_municipio in cols_disp else ""
    )
    df_exp["Municipio Detectado IGN"] = df_filtrado.get("MUNICIPIO_DETECTADO_IGN", "")
    df_exp["Provincia Detectada IGN"] = df_filtrado.get("PROVINCIA_DETECTADA_IGN", "")
    df_exp["Estado de Auditoría"] = df_filtrado.get("ESTADO_AUDITORIA", "")
    df_exp["Metros al Límite"] = df_filtrado.get("METROS_AL_LIMITE", np.nan)
    df_exp["Alerta Interprovincial"] = df_filtrado.get("ALERTA_INTERPROVINCIAL", "")
    df_exp["Alerta Coordenadas"] = df_filtrado.get("ALERTA_COORDENADAS", "")

    return df_exp.reset_index(drop=True)


def exportar_excel_en_memoria(
    df_export: pd.DataFrame,
    df_resumen: Optional[pd.DataFrame] = None,
) -> bytes:
    """Genera un archivo .xlsx en memoria con formato limpio y autoajuste de ancho."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Discrepancias")
        if df_resumen is not None and not df_resumen.empty:
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen_Auditoria")

        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(color="FFFFFF", bold=True)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_idx, col_cells in enumerate(ws.iter_cols(min_row=1), start=1):
                header_cell = col_cells[0]
                header_cell.fill = header_fill
                header_cell.font = header_font
                header_cell.alignment = Alignment(horizontal="center", vertical="center")

                max_len = 10
                for cell in col_cells[:500]:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_len:
                        max_len = len(val)
                ws.column_dimensions[get_column_letter(col_idx)].width = min(
                    max_len + 4, 48
                )

    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# INTERFAZ DE USUARIO STREAMLIT
# ==============================================================================


def _leer_hojas_excel(archivo_bytes: bytes) -> List[str]:
    xl = pd.ExcelFile(io.BytesIO(archivo_bytes), engine="openpyxl")
    return list(xl.sheet_names)


def seleccionar_hoja_por_defecto(archivo_bytes: bytes, hojas: List[str]) -> int:
    """
    Elige automáticamente el índice de la hoja más probable con el catálogo
    (evitando portadas o índices iniciales) al puntuar columnas clave detectadas.
    """
    if len(hojas) <= 1:
        return 0

    mejor_idx = 0
    mejor_score = -1

    try:
        with pd.ExcelFile(io.BytesIO(archivo_bytes), engine="openpyxl") as xl:
            for idx, hoja in enumerate(hojas):
                try:
                    df_muestra = pd.read_excel(xl, sheet_name=hoja, nrows=50)
                except Exception:
                    continue
                df_muestra = df_muestra.dropna(how="all")
                if df_muestra.empty:
                    continue

                det = detectar_columnas_inteligente(df_muestra)
                score = 0
                if det.get("latitud"):
                    score += 3
                if det.get("longitud"):
                    score += 3
                if det.get("municipio"):
                    score += 3
                if det.get("recurso"):
                    score += 1
                if (
                    det.get("latitud") is None or det.get("longitud") is None
                ) and detectar_columna_coordenadas_combinadas(df_muestra):
                    score += 4

                if score > mejor_score:
                    mejor_score = score
                    mejor_idx = idx
    except Exception:
        return 0

    return mejor_idx


def _leer_dataframe_excel(archivo_bytes: bytes, nombre_hoja: str) -> pd.DataFrame:
    df = pd.read_excel(
        io.BytesIO(archivo_bytes),
        sheet_name=nombre_hoja,
        engine="openpyxl",
    )
    # Limpiar nombres de columnas duplicados o no-string garantizando unicidad absoluta
    nuevas_cols: List[str] = []
    asignados: set[str] = set()
    vistas: Dict[str, int] = {}
    for i, c in enumerate(df.columns):
        base = str(c).strip() if pd.notna(c) and str(c).strip() else f"Columna_{i + 1}"
        if base not in asignados:
            nombre = base
            vistas[base] = 1
        else:
            contador = vistas.get(base, 1) + 1
            while f"{base}_{contador}" in asignados:
                contador += 1
            vistas[base] = contador
            nombre = f"{base}_{contador}"
        asignados.add(nombre)
        nuevas_cols.append(nombre)
    df.columns = nuevas_cols
    # Eliminar filas completamente vacías
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def _preparar_vista_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza columnas object con tipos Python mixtos para evitar errores Arrow en st.dataframe."""
    df_vista = df.copy()
    for col in df_vista.columns:
        serie = df_vista[col]
        if pd.api.types.is_object_dtype(serie):
            tipos = {type(v) for v in serie.dropna()}
            if len(tipos) > 1:
                df_vista[col] = serie.where(serie.isna(), serie.astype(str))
    return df_vista


def main() -> None:
    st.set_page_config(
        page_title="Auditoría Geoespacial IGN",
        page_icon="🗺️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("🗺️ Auditoría Geoespacial y Validación Municipal (IGN)")
    st.caption(
        "Contrasta las coordenadas GPS de cualquier catálogo Excel frente a los límites "
        "municipales oficiales del Instituto Geográfico Nacional (IGN)."
    )

    # 1. Cargar cartografía oficial del IGN (cacheada en memoria)
    with st.spinner("Cargando cartografía municipal oficial del IGN..."):
        try:
            gdf_espana_4326 = cargar_cartografia_ign()
        except Exception as exc:
            st.error(
                f"No se pudo descargar la cartografía oficial del IGN desde mapSpain: {exc}"
            )
            if st.button("🔄 Reintentar descarga del IGN"):
                st.cache_resource.clear()
                st.rerun()
            st.stop()

    # 2. Barra lateral: subida de archivo y parámetros
    with st.sidebar:
        st.header("1. Subir Catálogo Excel")
        archivo_subido = st.file_uploader(
            "Selecciona tu archivo (.xlsx)",
            type=["xlsx", "xlsm"],
            help="Sube cualquier archivo Excel con municipios y coordenadas GPS.",
        )

        st.divider()
        st.header("2. Ámbito Territorial")

        opciones_prov = [
            (cod, f"{cod} - {nom}") for cod, nom in sorted(PROVINCIAS_INE.items())
        ] + [("00", "00 - Toda España (sin filtro provincial)")]

        etiquetas_prov = [etq for _, etq in opciones_prov]
        codigos_prov = [cod for cod, _ in opciones_prov]
        idx_caceres = codigos_prov.index("10") if "10" in codigos_prov else 0

        sel_prov_etq = st.selectbox(
            "Provincia auditada",
            options=etiquetas_prov,
            index=idx_caceres,
            help="Selecciona manualmente la provincia de referencia para resolver homónimos y alertas interprovinciales.",
        )
        codigo_prov_seleccionado = codigos_prov[etiquetas_prov.index(sel_prov_etq)]

        st.divider()
        st.header("3. Escalas de Distancia (metros)")
        umbral_fronterizo = st.number_input(
            "Fronterizo inmediato (≤ m)",
            min_value=10,
            max_value=2000,
            value=100,
            step=10,
        )
        umbral_borde = st.number_input(
            "Borde cercano (≤ m)",
            min_value=int(umbral_fronterizo) + 10,
            max_value=5000,
            value=max(500, int(umbral_fronterizo) + 100),
            step=50,
        )
        umbral_moderado = st.number_input(
            "Desvío moderado (≤ m)",
            min_value=int(umbral_borde) + 50,
            max_value=20000,
            value=max(2000, int(umbral_borde) + 500),
            step=100,
        )
        umbral_notable = st.number_input(
            "Desvío notable (≤ m)",
            min_value=int(umbral_moderado) + 100,
            max_value=100000,
            value=max(6000, int(umbral_moderado) + 1000),
            step=500,
        )
        st.caption(f"Más de {umbral_notable:,} m → **ERROR_GRAVE**".replace(",", "."))

        st.divider()
        st.header("4. Limpieza de Coordenadas")
        reparar_sin_punto = st.checkbox(
            "Reparar coordenadas sin punto decimal (ej. 40136215 → 40.136215)",
            value=True,
            help="Corrige coordenadas a las que Excel en español eliminó el punto decimal al interpretarlo como separador de miles.",
        )

    # Estado inicial sin archivo
    if archivo_subido is None:
        st.info(
            "👈 **Sube un archivo Excel en el panel izquierdo** para iniciar la auditoría.\n\n"
            "**¿Qué hará la herramienta automáticamente?**\n"
            "1. Detectará los nombres de las columnas (Municipio, Latitud, Longitud, Nombre/Recurso e ID) aunque usen otros títulos.\n"
            "2. Validará cada coordenada contra los ~8.400 términos municipales oficiales del IGN (incluyendo alertas interprovinciales, Canarias y artículos invertidos como *ALDEA DEL OBISPO (LA)*).\n"
            "3. Generará un informe ejecutivo de discrepancias, el mapa interactivo anti-solapamiento y un Excel descargable con las incidencias."
        )
        st.stop()

    # Lectura del archivo Excel y selección de hoja
    archivo_bytes = archivo_subido.getvalue()
    try:
        hojas = _leer_hojas_excel(archivo_bytes)
    except Exception as exc:
        st.error(f"Error al abrir el archivo Excel: {exc}")
        st.stop()

    hoja_seleccionada = hojas[0]
    if len(hojas) > 1:
        idx_hoja_defecto = seleccionar_hoja_por_defecto(archivo_bytes, hojas)
        hoja_seleccionada = st.sidebar.selectbox(
            "Hoja del Excel",
            options=hojas,
            index=idx_hoja_defecto,
        )

    try:
        df_raw = _leer_dataframe_excel(archivo_bytes, hoja_seleccionada)
    except Exception as exc:
        st.error(f"Error al leer la hoja '{hoja_seleccionada}': {exc}")
        st.stop()

    if df_raw.empty:
        st.warning("La hoja seleccionada está vacía.")
        st.stop()

    # Detección flexible de columnas (incluye soporte para coordenadas combinadas)
    set_munis_ign = obtener_set_municipios_ign(gdf_espana_4326)
    cols_detectadas = detectar_columnas_inteligente(
        df_raw, municipios_ign_norm=set_munis_ign
    )

    df_trabajo = df_raw.copy()
    aviso_combinada = None
    if cols_detectadas["latitud"] is None or cols_detectadas["longitud"] is None:
        col_comb = detectar_columna_coordenadas_combinadas(df_trabajo)
        if col_comb:
            df_trabajo, col_lat_ext, col_lon_ext = separar_coordenadas_combinadas(
                df_trabajo, col_comb
            )
            cols_detectadas["latitud"] = col_lat_ext
            cols_detectadas["longitud"] = col_lon_ext
            aviso_combinada = col_comb

    columnas_df = list(df_trabajo.columns)
    columnas_con_ninguna = [SIN_COLUMNA_OPCION] + columnas_df

    def _idx_en_lista(lista: List[str], valor: Optional[str], default: int = 0) -> int:
        if valor in lista:
            return lista.index(valor)
        return default

    # Panel desplegable de mapeo de columnas (pre-rellenado automáticamente)
    with st.expander("⚙️ Mapeo automático de columnas (editable si fuera necesario)", expanded=False):
        if aviso_combinada:
            st.info(
                f"Se ha detectado una columna combinada de coordenadas (`{aviso_combinada}`) "
                "y se ha separado automáticamente en Latitud y Longitud."
            )
        c1, c2, c3, c4, c5 = st.columns(5)
        col_municipio = c1.selectbox(
            "Municipio *",
            options=columnas_df,
            index=_idx_en_lista(columnas_df, cols_detectadas["municipio"], 0),
        )
        col_latitud = c2.selectbox(
            "Latitud (Y) *",
            options=columnas_df,
            index=_idx_en_lista(columnas_df, cols_detectadas["latitud"], 0),
        )
        col_longitud = c3.selectbox(
            "Longitud (X) *",
            options=columnas_df,
            index=_idx_en_lista(
                columnas_df, cols_detectadas["longitud"], min(1, len(columnas_df) - 1)
            ),
        )
        col_recurso_sel = c4.selectbox(
            "Nombre / Recurso",
            options=columnas_con_ninguna,
            index=_idx_en_lista(columnas_con_ninguna, cols_detectadas["recurso"], 0),
        )
        col_id_sel = c5.selectbox(
            "ID (opcional)",
            options=columnas_con_ninguna,
            index=_idx_en_lista(columnas_con_ninguna, cols_detectadas["id"], 0),
        )

        col_recurso = None if col_recurso_sel == SIN_COLUMNA_OPCION else col_recurso_sel
        col_id = None if col_id_sel == SIN_COLUMNA_OPCION else col_id_sel

    if col_latitud == col_longitud:
        st.error(
            "Las columnas de **Latitud** y **Longitud** no pueden ser la misma. "
            "Ajusta el mapeo de columnas arriba."
        )
        st.stop()

    # Resolver provincia final (selección manual)
    if codigo_prov_seleccionado == "00":
        codigo_prov_efectivo = None
        nombre_prov_efectivo = "Toda España"
        st.info("🇪🇸 Ámbito territorial seleccionado: **Toda España (sin filtro provincial)**.")
    else:
        codigo_prov_efectivo = codigo_prov_seleccionado
        nombre_prov_efectivo = PROVINCIAS_INE.get(
            codigo_prov_seleccionado, codigo_prov_seleccionado
        )
        st.success(
            f"📍 Provincia auditada: **{nombre_prov_efectivo} (INE {codigo_prov_efectivo})**"
        )

    umbrales = {
        "fronterizo": float(umbral_fronterizo),
        "borde": float(umbral_borde),
        "moderado": float(umbral_moderado),
        "notable": float(umbral_notable),
    }

    # Ejecutar auditoría espacial
    with st.spinner("Ejecutando cruce espacial y calculando distancias frente al IGN..."):
        df_auditado = ejecutar_auditoria(
            df_origen=df_trabajo,
            gdf_espana_4326=gdf_espana_4326,
            col_municipio=col_municipio,
            col_latitud=col_latitud,
            col_longitud=col_longitud,
            col_recurso=col_recurso,
            col_id=col_id,
            codigo_provincia_ine=codigo_prov_efectivo,
            umbrales=umbrales,
            reparar_sin_punto=reparar_sin_punto,
        )

    reparadas_count = (
        int(df_auditado["COORD_REPARADA"].sum())
        if "COORD_REPARADA" in df_auditado.columns
        else 0
    )
    if reparadas_count > 0:
        st.info(
            f"🔧 Se normalizaron automáticamente **{reparadas_count}** registro(s) cuyas coordenadas "
            "venían sin punto decimal en el Excel (ej. `40136215` → `40.136215`)."
        )

    # Métricas del informe ejecutivo
    total_registros = len(df_auditado)
    sin_coords = df_auditado[df_auditado["ESTADO_AUDITORIA"] == "SIN_COORDENADAS"]
    correctos = df_auditado[df_auditado["ESTADO_AUDITORIA"] == "CORRECTO"]
    incidencias = df_auditado[
        ~df_auditado["ESTADO_AUDITORIA"].isin(["CORRECTO", "SIN_COORDENADAS"])
    ]
    interprovinciales = df_auditado[
        df_auditado["ALERTA_INTERPROVINCIAL"].str.contains("OTRA PROVINCIA", na=False)
    ]
    posibles_inversiones = df_auditado[
        df_auditado["ALERTA_COORDENADAS"] == "Posible inversión Lat/Lon"
    ]
    fuera_rango_esp = df_auditado[
        df_auditado["ALERTA_COORDENADAS"] == "Fuera de rango España"
    ]

    vali_count = total_registros - len(sin_coords)
    pct_correcto = (len(correctos) / vali_count * 100.0) if vali_count > 0 else 0.0

    # Resumen por categorías
    resumen_cat = (
        df_auditado["ESTADO_AUDITORIA"]
        .value_counts()
        .reset_index()
    )
    resumen_cat.columns = ["Categoría de Auditoría", "Nº Registros"]
    resumen_cat["% sobre Total"] = (
        resumen_cat["Nº Registros"] / max(total_registros, 1) * 100.0
    ).round(1)

    # Tarjetas KPI principales
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total registros", f"{total_registros:,}".replace(",", "."))
    k2.metric("Coords válidas", f"{vali_count:,}".replace(",", "."))
    k3.metric("Sin coordenadas", f"{len(sin_coords):,}".replace(",", "."))
    k4.metric("Correctos", f"{len(correctos):,}".replace(",", "."), f"{pct_correcto:.1f}%")
    k5.metric("Discrepancias", f"{len(incidencias):,}".replace(",", "."))
    k6.metric("Otra provincia", f"{len(interprovinciales):,}".replace(",", "."))

    # Alertas clave
    if not posibles_inversiones.empty:
        st.warning(
            f"⚠️ **Alerta de coordenadas**: Se han detectado **{len(posibles_inversiones)}** "
            "registros con posible inversión Latitud/Longitud (no se modifican automáticamente)."
        )
    if not fuera_rango_esp.empty:
        st.warning(
            f"⚠️ Se han detectado **{len(fuera_rango_esp)}** registros con coordenadas fuera "
            "del ámbito geográfico de España."
        )
    if len(incidencias) == 0 and len(sin_coords) == 0:
        st.success(
            "🎉 **¡Catálogo auditado sin discrepancias!** Todos los registros coinciden "
            "con el término municipal oficial del IGN."
        )

    # Pestañas principales
    tab_informe, tab_mapa, tab_repetidas, tab_completo = st.tabs(
        [
            f"📊 Informe y Discrepancias ({len(incidencias)})",
            "🗺️ Mapa Interactivo",
            "📍 Coordenadas Repetidas",
            f"📋 Catálogo Completo ({total_registros})",
        ]
    )

    # ------------------------------------------------------------------
    # PESTAÑA 1: INFORME EJECUTIVO, FILTROS Y EXPORTACIÓN EXCEL
    # ------------------------------------------------------------------
    with tab_informe:
        st.subheader(
            f"Informe Ejecutivo de Calidad — {nombre_prov_efectivo.upper()}"
        )

        col_tbl, col_chart = st.columns([1.1, 1.4])
        with col_tbl:
            st.dataframe(resumen_cat, width="stretch", hide_index=True)
        with col_chart:
            st.bar_chart(
                resumen_cat.set_index("Categoría de Auditoría")["Nº Registros"],
                height=260,
            )

        st.divider()
        st.subheader("Filtros de Discrepancias y Exportación a Excel")

        categorias_presentes = [
            c
            for c in PRIORIDAD_ESTADOS + ["SIN_COORDENADAS"]
            if c in df_auditado["ESTADO_AUDITORIA"].values
        ]
        categorias_defecto = [
            c
            for c in categorias_presentes
            if c not in ("CORRECTO", "SIN_COORDENADAS")
        ]

        f1, f2, f3 = st.columns([2.2, 1.0, 1.3])
        categorias_sel = f1.multiselect(
            "Categorías a mostrar y exportar",
            options=categorias_presentes,
            default=categorias_defecto,
            help="Puedes añadir 'SIN_COORDENADAS' o 'CORRECTO' si deseas incluirlos en el Excel.",
        )
        solo_otra_prov = f2.checkbox(
            "Solo alertas de otra provincia",
            value=False,
            disabled=len(interprovinciales) == 0,
        )
        busqueda_txt = f3.text_input(
            "Buscar por recurso o municipio",
            value="",
            placeholder="Ej. Plasencia, hotel...",
        )

        df_filtrado = df_auditado[
            df_auditado["ESTADO_AUDITORIA"].isin(categorias_sel)
        ].copy()

        if solo_otra_prov:
            df_filtrado = df_filtrado[
                df_filtrado["ALERTA_INTERPROVINCIAL"].str.contains(
                    "OTRA PROVINCIA", na=False
                )
            ]

        if busqueda_txt.strip():
            q = normalizar_texto(busqueda_txt)
            cols_busq = [
                c
                for c in [col_recurso, col_municipio, "MUNICIPIO_DETECTADO_IGN"]
                if c and c in df_filtrado.columns
            ]
            mask_busq = pd.Series(False, index=df_filtrado.index)
            for c in cols_busq:
                mask_busq = mask_busq | df_filtrado[c].apply(
                    lambda v: q in normalizar_texto(v)
                )
            df_filtrado = df_filtrado[mask_busq]

        # Vista de tabla de discrepancias
        cols_vista = []
        if col_id and col_id in df_filtrado.columns:
            cols_vista.append(col_id)
        if col_recurso and col_recurso in df_filtrado.columns:
            cols_vista.append(col_recurso)
        cols_vista.extend(
            [
                col_municipio,
                "MUNICIPIO_DETECTADO_IGN",
                "PROVINCIA_DETECTADA_IGN",
                "METROS_AL_LIMITE",
                "ESTADO_AUDITORIA",
                "ALERTA_INTERPROVINCIAL",
                col_latitud,
                col_longitud,
            ]
        )
        cols_vista = [c for c in cols_vista if c in df_filtrado.columns]

        st.write(
            f"Mostrando **{len(df_filtrado)}** registros según los filtros activos."
        )
        st.dataframe(
            _preparar_vista_dataframe(df_filtrado[cols_vista]),
            width="stretch",
            hide_index=True,
        )

        # Exportación Excel
        incluir_todas_cols = st.checkbox(
            "Incluir todas las columnas originales del Excel en la exportación",
            value=False,
            help="Si se desmarca, exporta el formato ejecutivo compacto igual que el notebook.",
        )

        df_export = construir_dataframe_exportacion(
            df_filtrado=df_filtrado,
            col_municipio=col_municipio,
            col_latitud=col_latitud,
            col_longitud=col_longitud,
            col_recurso=col_recurso,
            col_id=col_id,
            cols_extra={
                "descripcion": cols_detectadas.get("descripcion"),
                "direccion": cols_detectadas.get("direccion"),
                "url": cols_detectadas.get("url"),
            },
            incluir_todas_columnas=incluir_todas_cols,
        )

        nombre_archivo_prov = re.sub(
            r"[^A-Za-z0-9_-]+", "_", normalizar_texto(nombre_prov_efectivo).title()
        ).strip("_")
        nombre_excel_salida = f"Discrepancias_Catalogo_{nombre_archivo_prov or 'Espana'}.xlsx"

        excel_bytes = exportar_excel_en_memoria(df_export, df_resumen=resumen_cat)

        st.download_button(
            label=f"📥 Descargar Excel de Discrepancias ({len(df_export)} registros)",
            data=excel_bytes,
            file_name=nombre_excel_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            disabled=df_export.empty,
        )

    # ------------------------------------------------------------------
    # PESTAÑA 2: MAPA INTERACTIVO UNIFICADO (ANTI-SOLAPAMIENTO)
    # ------------------------------------------------------------------
    with tab_mapa:
        m_col1, m_col2 = st.columns([2, 1])
        mostrar_correctos_mapa = m_col1.checkbox(
            "Mostrar también puntos CORRECTOS (🔵) en el mapa",
            value=True,
            help="Desactívalo si el archivo es muy grande y solo quieres revisar incidencias.",
        )

        with st.spinner("Generando mapa interactivo consolidado..."):
            mapa_folium, excluidos_mapa = generar_mapa_folium(
                df_auditado=df_auditado,
                gdf_espana_4326=gdf_espana_4326,
                col_municipio=col_municipio,
                col_latitud=col_latitud,
                col_longitud=col_longitud,
                col_recurso=col_recurso,
                col_id=col_id,
                codigo_provincia_ine=codigo_prov_efectivo,
                nombre_provincia=nombre_prov_efectivo,
                umbrales=umbrales,
                mostrar_correctos=mostrar_correctos_mapa,
            )
            mapa_html = mapa_folium.get_root().render()

        m_col2.download_button(
            label="💾 Descargar mapa interactivo (.html)",
            data=mapa_html.encode("utf-8"),
            file_name=f"mapa_auditoria_{nombre_archivo_prov or 'Espana'}.html",
            mime="text/html",
            width="stretch",
        )

        if excluidos_mapa > 0:
            st.info(
                f"ℹ️ {excluidos_mapa} punto(s) con coordenadas fuera de España no se "
                "muestran en el mapa para mantener el encuadre geográfico."
            )

        if hasattr(st, "iframe"):
            st.iframe(mapa_html, height=660)
        else:
            import streamlit.components.v1 as components

            components.html(mapa_html, height=660, scrolling=False)

    # ------------------------------------------------------------------
    # PESTAÑA 3: INSPECCIÓN DE COORDENADAS REPETIDAS
    # ------------------------------------------------------------------
    with tab_repetidas:
        st.subheader("Inspección Rápida de Discrepancias y Coordenadas Repetidas")
        solo_inc_rep = st.checkbox(
            "Mostrar únicamente registros con discrepancias (excluir CORRECTO)",
            value=True,
        )
        df_rep = construir_tabla_coordenadas_repetidas(
            df_auditado=df_auditado,
            col_municipio=col_municipio,
            col_latitud=col_latitud,
            col_longitud=col_longitud,
            col_recurso=col_recurso,
            col_id=col_id,
            solo_incidencias=solo_inc_rep,
        )
        if df_rep.empty:
            st.info("No hay registros con coordenadas válidas bajo el filtro seleccionado.")
        else:
            st.dataframe(_preparar_vista_dataframe(df_rep), width="stretch", hide_index=True)

    # ------------------------------------------------------------------
    # PESTAÑA 4: CATÁLOGO COMPLETO AUDITADO
    # ------------------------------------------------------------------
    with tab_completo:
        st.subheader("Catálogo Completo con Columnas de Auditoría IGN")
        st.dataframe(_preparar_vista_dataframe(df_auditado), width="stretch", hide_index=True)
        excel_completo_bytes = exportar_excel_en_memoria(
            df_auditado, df_resumen=resumen_cat
        )
        st.download_button(
            label=f"📥 Descargar Catálogo Completo Auditado ({total_registros} registros)",
            data=excel_completo_bytes,
            file_name=f"Catalogo_{nombre_archivo_prov or 'Espana'}_AUDITADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
