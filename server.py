"""juris-pj-mcp: servidor MCP para la Jurisprudencia Nacional Sistematizada del Poder Judicial del Peru (jurisprudencia.pj.gob.pe).

Busca resoluciones (texto, nivel, distrito, especialidad, sala, pretension, palabras clave, expediente, tipo de recurso/resolucion, anio), lee las fichas/resumenes con sumilla COMPLETA y palabras clave, obtiene la URL publica del PDF de cada resolucion, exporta la busqueda a Excel, consulta los laudos arbitrales (texto, anio, expediente, palabra clave y orden/forma) y los autocompletes del buscador desde tu harness MCP.

Conexion directa, sin navegador y sin credenciales: todo el contenido consultado es publico. El servidor nunca guarda archivos ni trunca datos: devuelve datos limpios y URLs de descarga (solo el Excel viene como base64 con cap declarado).

Proyecto NO oficial, sin afiliacion con el Poder Judicial ni el Estado peruano.
"""

import asyncio
import base64
import html as html_mod
import json
import logging
import os
import re
import time
import unicodedata
import uuid
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

import httpx
from mcp.server.fastmcp import FastMCP

logging.getLogger("httpx").setLevel(logging.WARNING)

BASE = "https://jurisprudencia.pj.gob.pe"
APP = BASE + "/jurisprudenciaweb"
PAGE_INICIO = APP + "/faces/page/inicio.xhtml"
PAGE_COMERCIALES = BASE + "/jurisprudenciaweb/faces/page/comerciales.xhtml"
PAGE_RESULTADO = BASE + "/jurisprudenciaweb/faces/page/resultado.xhtml"
SERVE_PDF = APP + "/ServletDescarga"

TIMEOUT = float(os.getenv("PJ_TIMEOUT", "240"))
TOOL_TIMEOUT = float(os.getenv("PJ_TOOL_TIMEOUT", "300"))
MAX_CONCURRENCY = int(os.getenv("PJ_MAX_CONCURRENCY", "2"))
INFLIGHT_CAP = int(os.getenv("PJ_INFLIGHT_CAP", "8"))
SESION_TTL = int(os.getenv("PJ_SESION_TTL", "900"))
SESION_MAX = int(os.getenv("PJ_SESION_MAX", "10"))
EXCEL_MAX_BYTES = int(os.getenv("PJ_EXCEL_MAX_MB", "12")) * 1024 * 1024
VERIFY_TLS = os.getenv("PJ_VERIFY", "1").strip() not in ("0", "false", "no")
MAX_PAGINAS = 6
TAM_PAGINA = 10  # el sitio pagina a 10 filas por vista (verificado 2026-10-15)

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"

HDRS = {
    "User-Agent": UA,
    "Accept-Language": "es-PE,es;q=0.9",
    "Origin": "https://jurisprudencia.pj.gob.pe",
}

CT_FORM = "application/x-www-form-urlencoded; charset=UTF-8"

NIVEL_CODES: dict[str, str] = {"1": "Corte Suprema", "2": "Corte Superior"}

ESPECIALIDADES: dict[str, str] = {
    "1": "Civil", "2": "Comercial", "3": "Contencioso Administrativo",
    "4": "Familia Civil", "5": "Familia Tutelar", "6": "Familia Penal",
    "7": "Contencioso Adm. Laboral", "8": "Contencioso Adm. Previsional",
    "9": "Laboral", "10": "Penal", "11": "Revision de Procedimiento Coactivo",
    "12": "Constitucional",
}

DISTRITOS: dict[str, str] = {
    "1": "Amazonas", "2": "Ancash", "3": "Apurimac", "4": "Arequipa",
    "5": "Ayacucho", "6": "Cajamarca", "7": "Callao", "8": "Canete",
    "9": "Lima Norte", "10": "Cusco", "11": "Huancavelica", "12": "Huanuco",
    "13": "Huaura", "14": "Ica", "15": "Junin", "16": "La Libertad",
    "17": "Lambayeque", "18": "Lima", "19": "Loreto", "20": "Piura",
    "21": "Puno", "22": "San Martin", "23": "Tacna", "24": "Ucayali",
    "25": "Del Santa", "26": "Tumbes", "27": "Madre de Dios",
    "28": "Moquegua", "29": "Pasco", "30": "Lima Sur", "31": "Sullana",
    "38": "Lima Este", "39": "Puente Piedra - Ventanilla",
    "41": "Selva Central",
}

TIPOS_RECURSO: dict[str, str] = {
    "1": "Revision de Sentencia", "2": "Casacion", "3": "Recurso de Nulidad",
    "4": "Consulta", "5": "Extradiccion Activa",
    "6": "Recurso de Queja Excepcional", "7": "Revision de Medida Disciplinaria",
    "8": "Amparo", "9": "Accion Popular", "10": "Inhibicion", "11": "Queja",
    "14": "Recurso de Queja Directa", "15": "Recurso de Queja Ordinaria",
    "16": "Casacion Previsional", "18": "Extradiccion Pasiva", "19": "Apelacion",
    "20": "Competencia", "21": "Queja NCPP", "22": "Recusacion",
    "23": "Querella", "24": "Nulidad", "27": "Medida Cautelar de Proceso de Amparo",
}

TIPOS_RESOLUCION: dict[str, str] = {
    "0": "Todos", "1": "Auto", "2": "Ejecutoria Suprema",
    "3": "Sentencia de Vista",
}

ORDEN_CODES: dict[str, str] = {
    "0": "0", "nro_de_recurso": "0", "nro_recurso": "0",
    "64": "64", "anio": "64",
    "66": "66", "pretension": "66", "pretension_delito": "66",
    "67": "67", "tipo_de_resolucion": "67", "tipo_resolucion": "67",
    "68": "68", "sala_suprema": "68",
    "21": "21", "fecha": "21", "fecha_resolucion": "21",
    "fecha_de_resolucion": "21",
}

ORDEN_ETIQUETAS = [("0", "Nro de Recurso"), ("64", "Anio"),
                   ("66", "Pretension/Delito"), ("67", "Tipo de Resolucion"),
                   ("68", "Sala Suprema"), ("21", "Fecha Resolucion")]

ANIOS = [str(a) for a in range(2026, 1997, -1)] + ["1982"]

CAMPOS_FILA = [
    (r"Especialidad", "especialidad"),
    (r"[ÓÓo]rgano Jurisdiccional", "organo_jurisdiccional"),
    (r"Sala Suprema", "sala_suprema"),
    (r"Norma de Derecho Interno", "norma_derecho_interno"),
    (r"Tipo Resoluci[óo]n", "tipo_resolucion"),
    (r"Fecha Resoluci[óo]n", "fecha_resolucion"),
    (r"Preten[cs]i[óo]n\s*/\s*Delito", "pretension_delito"),
    (r"Sumilla", "sumilla"),
    (r"Palabras Clave", "palabras_clave"),
]

RE_UUID = r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"

CODIGOS_ORDEN = ("21", "0", "64", "66", "67", "68")

EVALUADOS = ("1", "true", "si", "s", "on", "yes")

ALIAS_TODOS = ("todos", "todo", "-- todos --", "all", "0")

_REPO_GITHUB = "pipaacebedo/juris-pj-mcp"
_VERSION = "0.0.0"
try:
    _VERSION = str(json.loads(Path(__file__).resolve().parent.joinpath("manifest.json").read_text(encoding="utf-8")).get("version", _VERSION))
except Exception:
    pass


def _semver_mayor(remota: str, instalada: str) -> bool:
    def partes(v: str) -> list[int]:
        return [int(x) for x in re.findall(r"\d+", v)[:3]] or [0]

    return partes(remota) > partes(instalada)


def _update_remota() -> str:
    """tag semver de un release mas nuevo que el instalado; '' si no hay, si esta desactivado (PJ_UPDATE_CHECK=0) o si el chequeo falla (NUNCA lanza)."""
    if os.getenv("PJ_UPDATE_CHECK", "1").strip() in ("0", "false", "no"):
        return ""
    try:
        r = httpx.get(f"https://api.github.com/repos/{_REPO_GITHUB}/releases?per_page=5", headers={"Accept": "application/vnd.github+json", "User-Agent": "juris-pj-mcp"}, timeout=3.0)
        for rel in (r.json() or []):
            tag = str(rel.get("tag_name", ""))
            if re.fullmatch(r"v?\d+\.\d+(\.\d+)?", tag) and _semver_mayor(tag, _VERSION):
                return tag.lstrip("v")
        return ""
    except Exception:
        return ""


def _linea_update() -> str:
    """parrafo UPDATE DISPONIBLE para las INSTRUCCIONES; '' si no hay version mas nueva."""
    remota = _update_remota()
    if not remota:
        return ""
    return f"UPDATE DISPONIBLE - existe la version v{remota} (instalada v{_VERSION}): pide al usuario descargar el .mcpb nuevo y reinstalarlo (https://github.com/{_REPO_GITHUB}/releases/latest); el estado de sesion del sitio no se traslada entre versiones.\n"


INSTRUCCIONES = (
    "Servidor de la JURISPRUDENCIA NACIONAL SISTEMATIZADA del Poder Judicial peruano (resoluciones sistematizadas por el PJ: casaciones, apelaciones, ejecutorias supremas, autos, sentencias de vista y laudos). Lee en 3 niveles:\n"
    "0 DESCUBRIR - buscar_resoluciones(texto='...') o con filtros: nivel ('Corte Suprema' default / 'Corte Superior' + distrito + especialidad del distrito + sala EXACTA via listar_salas), pretension y palabra_clave (valores EXACTOS via los autocompletes), expediente, tipo_recurso/tipo_resolucion, anio, ncpp/nlpt (con especialidad Penal/Laboral), autos y orden/forma. Cada llamada trae 'ses', filas con 'index' GLOBAL (0-based y continuo entre paginas), 'total' y 'paginas'. La SUMILLA viene COMPLETA (sin truncar: es el texto de la decision).\n"
    "SEMANTICA: buscar_resoluciones con SOLO texto reproduce la busqueda del home del sitio (el conjunto mas amplio); al AGREGAR filtros, el sitio refina como en su UI (totales mas angostos); los filtros SOLOS (sin texto) son validos (el sitio demanda un criterio: el servidor aterriza con un token interno y el filtro queda puro). La busqueda de LAUDOS es aparte: buscar_laudos (sesion propia).\n"
    "1 DETALLE - ficha_resolucion(ses, index) y fichas_resoluciones(ses, indices, cap 50) para lotes (los fallidos van a 'faltas' con su motivo).\n"
    "2 DOCUMENTOS - url_resolucion(uuid) entrega la URL PUBLICA del PDF del PJ (descarga directa; el servidor no guarda archivos). exportar_excel(ses) exporta la busqueda a XLS en base64 (cap declarado de 12 MB; si lo supera, error con tamanos, NUNCA truncamiento silencioso). Cada respuesta trae 'ms' (milisegundos reales del server en esa llamada).\n"
    "RENDIMIENTO REAL: el sitio es LENTO (aterrizaje 9-18 s; POST de filtros 2-8 s; cada pagina paginada 8-12 s; PDF 4-17 s). El sitio SIRVE SIEMPRE 10 filas por pagina y 'paginas' (1..6) trae paginas consecutivas en UNA llamada: prefiere 1 llamada con paginas=3..6 a muchas llamadas de 1 pagina. Catalogos: listar_filtros (todas las listas) y listar_salas (~15-35 s la primera vez, cache 30 min). Toda tool responde dentro de PJ_TOOL_TIMEOUT (300 s): si el sitio congela el socket, la llamada vuelve con error y pista, nunca cuelga.\n"
    "Politica de sesion: 'ses' GLOBAL compartido por las tools de la MISMA busqueda, TTL de 15 min DESLIZANTE desde la ultima actividad, eliminacion LRU. Cambio de filtros con el mismo 'ses' re-ejecuta la busqueda conservando el id; un 'ses' desconocido (vencido) junto con filtros lanza una busqueda NUEVA con id nuevo y lo declara en la 'nota'. 'ses' vencido con pagina>1 -> error 'sesion expirada o no encontrada': repite la busqueda. Si la respuesta trae 'error', usa la 'pista' y repite.\n"
)

INSTRUCCIONES += _linea_update()

mcp = FastMCP("Jurisprudencia Nacional Sistematizada (No oficial)", instructions=INSTRUCCIONES, stateless_http=True)

_SEM: asyncio.Semaphore | None = None
_INFLIGHT = 0


def _sem() -> asyncio.Semaphore:
    global _SEM
    if _SEM is None:
        _SEM = asyncio.Semaphore(max(1, MAX_CONCURRENCY))
    return _SEM


# ------------------------------------------------------------------ helpers

def _sin_acentos(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def _plain(s: str | None) -> str:
    return _sin_acentos(s or "").strip().lower()


def _limpia(s: str | None) -> str:
    """limpia HTML/NBSP/mojibake SIN truncar el contenido (regla del server)."""
    t = re.sub(r"(?is)<script[\s\S]*?</script>", " ", s or "")
    t = re.sub(r"<[^>]+>", " ", t)
    t = html_mod.unescape(t)
    t = t.replace("\xa0", " ").replace("\u200b", "")
    t = t.replace("Â°", "°").replace("Ã±", "ñ").replace("Ã©", "é") \
        .replace("Ã­", "í").replace("Ã³", "ó").replace("Ãº", "ú") \
        .replace("Ã‘", "Ñ")
    return re.sub(r"\s+", " ", t).strip()


def _ceil(a: int, b: int) -> int:
    return -(-a // max(1, b))


async def con_watchdog(coro, nombre: str):
    """TODA tool responde dentro de TOOL_TIMEOUT: si la tarea se cuelga (socket sin FIN del sitio, etc), la cancela y devuelve un error explicando; NUNCA deja al cliente esperando indefinidamente."""
    try:
        return await asyncio.wait_for(coro, timeout=TOOL_TIMEOUT)
    except TimeoutError:
        raise ErrorSite(
            f"la operacion {nombre!r} no respondio en {TOOL_TIMEOUT:.0f}s (el sitio esta muy lento o la conexion quedo colgada); reintenta la llamada en unos minutos")


def _texto_busq(filtros: dict) -> str:
    """texto para la busqueda inicial; con solo-filtros usa un token interno ('zz') que el POST de filtros siempre reemplaza (el sitio borra el texto cuando el refino llega con txtBusqueda vacio: verificado 2026-10-15)."""
    t = _limpia(str(filtros.get("texto", "")))
    if t:
        return t
    pc = _limpia(str(filtros.get("palabra_clave", "")))
    exp = _limpia(str(filtros.get("expediente", "")))
    pre = _limpia(str(filtros.get("pretension", "")))
    return pc or exp or pre or "zz"


def _verdadero(v: Any) -> bool:
    return _plain(str(v or "")) in EVALUADOS


def _ms_desde(t0: float) -> int:
    """milisegundos transcurridos en el server (dato de latencia por llamada)."""
    return round((time.perf_counter() - t0) * 1000)


ORDEN_OPCIONES_TXT = "'Nro de Recurso', 'Anio', 'Pretension/Delito', 'Tipo de Resolucion', 'Sala Suprema' o 'Fecha Resolucion'"


def _orden_code(valor: Any) -> str | None:
    """resuelve 'orden' por etiqueta del catalogo ('Nro de Recurso', ...), alias o codigo; None si no es valido (no hay fallback silencioso)."""
    v = _plain(str(valor or ""))
    if not v:
        return None
    return ORDEN_CODES.get(re.sub(r"[^a-z0-9]+", "_", v).strip("_"))


def _forma_valor(valor: Any) -> str:
    return "ASC" if _plain(str(valor or "")) == "asc" else "DESC"


def _valida_orden_forma(filtros: dict[str, str]) -> None:
    o = str(filtros.get("orden", "")).strip()
    if o and _orden_code(o) is None:
        raise ErrorSite(f"orden {o!r} no es valido; usa {ORDEN_OPCIONES_TXT}; los mismos nombres salen en listar_filtros(categoria='orden')")
    f = str(filtros.get("forma", "")).strip()
    if f and _plain(f) not in ("asc", "desc"):
        raise ErrorSite(f"forma {f!r} no es valida; usa 'ASC' o 'DESC'")


def _etiqueta(valor: Any, cat: dict[str, str], pista: str) -> str:
    """resuelve un filtro de catalogo: PRIMERO match EXACTO del nombre (dos pasadas; evita que 'Penal' caiga en 'Familia Penal'), luego por substring. 'todos'/'0' equivalen a no filtrar (devuelve '')."""
    v = _plain(str(valor or ""))
    if not v:
        return ""
    if v in ALIAS_TODOS:
        return ""
    exactos = [cod for cod, nom in cat.items() if _plain(nom) == v]
    if exactos:
        return exactos[0]
    cercanos = [cod for cod, nom in cat.items() if v and v in _plain(nom)]
    if cercanos:
        return cercanos[0]
    raise ErrorSite(f"{pista}: {valor!r} no es valido; ejemplos: "
                    + "; ".join(list(dict.fromkeys(cat.values()))[:6]))


def _opciones(html: str, nombre: str) -> list[tuple[str, str]]:
    x = html_mod.unescape(html or "")
    m = re.search(r'name="formBuscador:' + nombre +
                  r'"[^>]*>([\s\S]{0,9000}?)</select>', x)
    if not m:
        m = re.search(r'name="formTributario:' + nombre +
                      r'"[^>]*>([\s\S]{0,9000}?)</select>', x)
    if not m:
        return []
    opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>',
                      m.group(1))
    return [(v, _limpia(l)) for v, l in opts if _limpia(l)]


def _match_combo(opciones: list[tuple[str, str]], valor: str) -> str | None:
    p = _plain(valor)
    for cod, _lab in opciones:
        if cod == valor:
            return valor
    for cod, lab in opciones:
        pl = _plain(lab)
        if pl == p or (p and p in pl):
            return cod
    return None


def _msgs(x: str) -> str:
    xs = re.findall(
        r'<span[^>]*class="rf-msgs-(?:sum|detailed)"[^>]*>([^<]*)</span>',
        x or "")
    if not xs:
        xs = re.findall(r'class="rf-msgs-err"[^>]*>[\s\S]{0,300}?</span>',
                        x or "")
        xs = [re.sub(r"<[^>]+>", " ", s) for s in xs]
    partes = []
    for s in xs[:4]:
        t = _limpia(html_mod.unescape(s))
        t = re.sub(r"^rf-msgs.*?:\s*", "", t)
        t = re.sub(r"^formBuscador\:\w+\:\s*", "", t)
        if t:
            partes.append(t[:140])
    return "; ".join(partes) if partes else ""


def _parsear_filas(x: str) -> dict[int, dict]:
    """filas del panel de resultados, SIN TRUNCAZIONES (la sumilla llega completa tal como la publica el sitio)."""
    x = html_mod.unescape(x or "")
    out: dict[int, dict] = {}
    ms = list(re.finditer(
        r'<div class="rf-p " id="formBuscador:repeat:(\d+):', x))
    for j, m in enumerate(ms):
        idx = int(m.group(1))
        if idx in out:
            continue
        seg = x[m.start(): ms[j + 1].start() if j + 1 < len(ms)
                else min(m.start() + 14000, len(x))]
        fila: dict[str, Any] = {"index": idx}
        mcab = re.search(r'<td>\s*(?:<input[^>]*/>)?\s*</td>\s*<td>\s*<span[^>]*>\s*([^<]*?)\s*</span></td>\s*<td>\s*<span[^>]*>\s*([^<]*?)\s*</span></td>', seg)
        if mcab:
            fila["recurso"] = _limpia(mcab.group(1))
            fila["expediente"] = _limpia(mcab.group(2))
        for patr, clave in CAMPOS_FILA:
            mval = re.search(patr + r'[\s\S]{0,50}?</div>\s*<div class="col-md-12">\s*([\s\S]{0,60000}?)\s*</div>', seg)
            if mval:
                # el label EXISTE en la fila: la clave va aunque el sitio la
                # publique vacia ("" = publicada vacia; clave ausente = el
                # sitio no publica ese campo en esta fila); nada se inventa
                fila[clave] = _limpia(mval.group(1))
        mu = re.search(r"ServletDescarga\?uuid=" + RE_UUID, seg)
        if mu:
            fila["uuid"] = mu.group(0).split("uuid=")[-1]
        if fila.get("expediente") or fila.get("sumilla") or fila.get("uuid"):
            fila.setdefault("recurso", "")
            fila.setdefault("expediente", "")
            out[idx] = fila
    return out


def _ordenar(rows: dict[int, dict]) -> list[dict]:
    return [rows[k] for k in sorted(rows)]


# --------------------------------------------------------------------- sesion

class ErrorSite(RuntimeError):
    """error elegible del flujo del sitio (mensaje en espanol)."""


class ErrorValidacion(ErrorSite):
    """el JSF del sitio rechazo la operacion (msg con los msgs del render)."""


class ErrorRango(ErrorSite):
    """error DETERMINISTA de rango (pagina/index): no se reintenta ni re-ejecuta."""


class Sesion:
    """Sesion del buscador: cliente httpx + estado JSF + filas por index."""

    def __init__(self) -> None:
        self.key = uuid.uuid4().hex[:8]
        self.ts = time.time()
        self.client = httpx.AsyncClient(
            timeout=httpx.Timeout(TIMEOUT), headers=HDRS,
            follow_redirects=False, verify=VERIFY_TLS)
        self.sid = ""
        self.vs = ""
        self.modo = ""
        self.filtros: dict | None = None
        self.filas: dict[int, dict] = {}
        self.total = 0
        self.paginas = 1
        self.aviso_paginas = ""
        self.corpus = 0
        self.html_resultado = ""
        self.html_rf = ""
        self.html_comercial = ""
        self.campos_filtro: dict | None = None
        self.lock = asyncio.Lock()

    async def close(self) -> None:
        try:
            await self.client.aclose()
        except Exception:
            pass

    def touch(self) -> None:
        self.ts = time.time()

    def _url(self, page: str) -> str:
        return page + (";jsessionid=" + self.sid if self.sid else "")

    @staticmethod
    def _norm(loc: str) -> str:
        return (loc or "").replace("http://jurisprudencia.pj.gob.pe",
                                   "https://jurisprudencia.pj.gob.pe")

    async def _get(self, url: str) -> httpx.Response:
        global _INFLIGHT
        async with _sem():
            while _INFLIGHT >= INFLIGHT_CAP:
                await asyncio.sleep(0.2)
            _INFLIGHT += 1
            try:
                try:
                    return await self.client.get(url)
                except (httpx.ReadTimeout, httpx.ConnectError):
                    await asyncio.sleep(2)
                    try:
                        return await self.client.get(url)
                    except (httpx.ReadTimeout, httpx.ConnectError) as e2:
                        raise ErrorSite(
                            f"el servidor del PJ no respondio ({type(e2).__name__}); es lento o se saturó: reintenta en unos segundos")
            finally:
                _INFLIGHT -= 1

    async def _post(self, url: str, data: dict[str, str], *,
                    ajax: bool = False,
                    referer: str = PAGE_RESULTADO) -> httpx.Response:
        global _INFLIGHT
        hb = dict(HDRS)
        hb["Referer"] = referer
        hb["Content-Type"] = CT_FORM
        if ajax:
            hb["Faces-Request"] = "partial/ajax"
        async with _sem():
            while _INFLIGHT >= INFLIGHT_CAP:
                await asyncio.sleep(0.2)
            _INFLIGHT += 1
            try:
                try:
                    return await self.client.post(
                        url, content=urlencode(data).encode("utf-8"), headers=hb)
                except (httpx.ReadTimeout, httpx.ConnectError):
                    await asyncio.sleep(2)
                    try:
                        return await self.client.post(
                            url, content=urlencode(data).encode("utf-8"), headers=hb)
                    except (httpx.ReadTimeout, httpx.ConnectError) as e2:
                        raise ErrorSite(
                            f"el servidor del PJ no respondio ({type(e2).__name__}); es lento o se saturó: reintenta en unos segundos")
            finally:
                _INFLIGHT -= 1

    async def _abrir_inicio(self) -> str:
        r = await self._get(PAGE_INICIO)
        if r.status_code != 200:
            raise ErrorSite(f"inicio no disponible (HTTP {r.status_code})")
        x = r.text
        m = re.search(r'action="[^"]*;jsessionid=([^"?]+)', x)
        self.sid = m.group(1) if m else ""
        self.vs = self._vs_de(x)
        if not self.vs:
            raise ErrorSite("inicio sin ViewState; el sitio puede haber cambiado")
        return x

    @staticmethod
    def _vs_de(x: str) -> str:
        mv = re.search(r'name="javax\.faces\.ViewState"[^>]*value="([^"]+)"', x)
        if mv:
            return mv.group(1)
        mv = re.search(r'<update id="javax\.faces\.ViewState"><!\[CDATA\[([^\]]+)\]', x or "")
        return mv.group(1) if mv else ""

    @staticmethod
    def _jsfcljs(x: str, form: str) -> list[list[tuple[str, str]]]:
        """los pares de 'mojarra.jsfcljs' de un form (ids j_idt dinamicos)."""
        xu = html_mod.unescape(x or "").replace("\\'", "'")
        out: list[list[tuple[str, str]]] = []
        pat = (r"mojarra\.jsfcljs\(document\.getElementById\('" + re.escape(form) +
               r"'\),\{([^}]*)\}")
        for m in re.finditer(pat, xu):
            pares = re.findall(r"'([^']+)'\s*:\s*'([^']*)'", m.group(1))
            if pares:
                out.append(pares)
        return out

    # ------------------------------------------------------------ busquedas

    async def buscar(self, modo: str, filtros: dict, pagina: int,
                     paginas: int) -> list[dict]:
        async with self.lock:
            return await self._buscar_impl(modo, filtros, pagina, paginas)

    async def _buscar_impl(self, modo: str, filtros: dict, pagina: int,
                           paginas: int) -> list[dict]:
        cristalizados = {k: str(v).strip() for k, v in filtros.items()
                         if str(v or "").strip()}
        continuar = (self.modo == modo and self.modo
                     and self.filtros is not None
                     and self._coinciden(cristalizados)
                     and (pagina > 1 or not cristalizados))
        intento_salva = False
        filas: list[dict] = []
        while True:
            try:
                if not continuar:
                    await self._nueva_busqueda(modo, cristalizados)
                else:
                    self.filtros = dict(self.filtros or {})
                    self.filtros.update(cristalizados)
                filas = []
                self.aviso_paginas = ""
                rango = range(pagina, pagina + max(1, paginas))
                for pag in rango:
                    tope = self.paginas if self.total else 0
                    if pag > 1 and pag > tope:
                        if pag > pagina:
                            if self.total:
                                self.aviso_paginas = f"la busqueda tiene {self.paginas} paginas"
                            break
                        raise ErrorRango(f"pagina {pag} fuera de rango: la busqueda tiene {self.paginas} paginas; revisa 'total' y 'paginas' de la ultima busqueda")
                    if pag == 1:
                        bloque = [f for f in _ordenar(self.filas)
                                  if f["index"] < TAM_PAGINA]
                    else:
                        bloque = await self._paginar(pag)
                    filas.extend(bloque)
                break
            except ErrorValidacion:
                raise
            except ErrorRango:
                raise
            except ErrorSite as e:
                if continuar and not intento_salva and self.filtros is not None:
                    intento_salva = True
                    continuar = False
                    cristalizados = dict(self.filtros)
                    cristalizados.pop("pagina", None)
                    cristalizados.pop("paginas", None)
                    cristalizados.pop("ses", None)
                    continue
                if self.total:
                    raise ErrorSite(
                        f"{e}; la busqueda sigue viva en la sesion {self.key}")
                raise
        self.touch()
        return filas

    def _coinciden(self, nuevos: dict[str, str]) -> bool:
        antiguos = {k: str(v).strip() for k, v in (self.filtros or {}).items()
                    if str(v or "").strip()
                    and k not in ("pagina", "paginas", "ses")}
        return all(antiguos.get(k) == v for k, v in nuevos.items())

    async def _nueva_busqueda(self, modo: str, filtros: dict[str, str]) -> None:
        _valida_orden_forma(filtros)          # falla RAPIDO, sin tocar la red
        self.modo = modo
        self.filas = {}
        self.html_resultado = ""
        self.html_rf = ""
        self.total = 0
        self.paginas = 1
        try:
            if modo == "laudos":
                await self._iniciar_laudos()
                await self._buscar_laudos(filtros)
            else:
                campos = self._campos_filtro(filtros)  # valida y resuelve sala
                self.campos_filtro = campos
                await self._iniciar_principal(filtros)
                # con solo texto/orden/forma/autos el landing (etapa 1) basta;
                # con filtros el sitio refina como en su UI (etapa 2 = POST)
                hay_especiales = any(
                    k not in ("texto", "orden", "forma", "autos")
                    for k in filtros)
                if campos["distrito"] not in ("", "0") or \
                        campos["sala"] not in ("", "0"):
                    await self._cadena_dependencias()
                if hay_especiales:
                    await self._aplicar_filtros(filtros)
            self.filtros = {k: str(v).strip() for k, v in filtros.items()
                            if k not in ("pagina", "paginas", "ses")
                            and str(v or "").strip()}
        except BaseException:
            # una busqueda fallida (error, watchdog o cancelacion) NO deja la
            # sesion a medio inicializar: queda limpia y con mensajes veraces
            self.modo = ""
            self.filtros = None
            self.campos_filtro = None
            self.filas = {}
            self.total = 0
            self.paginas = 1
            raise

    async def _iniciar_principal(self, filtros: dict[str, str]) -> None:
        x = await self._abrir_inicio()
        pares_busc = None
        for pares in self._jsfcljs(x, "formBuscador"):
            dic = dict(pares)
            if dic.get("forward") == "buscar" and dic.get("busqueda"):
                pares_busc = pares
                break
        if pares_busc is None:
            raise ErrorSite("el botón de busqueda del inicio no esta disponible; el sitio puede haber cambiado")
        dic = dict(pares_busc)
        btn = next(iter(dic))
        data: dict[str, str] = {"formBuscador": "formBuscador",
                                "javax.faces.ViewState": self.vs}
        for k, v in pares_busc:
            data[k] = v
        orden = _orden_code(filtros.get("orden"))
        forma = _forma_valor(filtros.get("forma"))
        orden_previo = None
        forma_previa = None
        for k, v in pares_busc:
            if k == btn:
                continue
            if v in CODIGOS_ORDEN:
                orden_previo = k
            elif v in ("ASC", "DESC"):
                forma_previa = k
        if orden_previo:
            data[orden_previo] = orden or "21"
        if forma_previa:
            data[forma_previa] = forma
        data["formBuscador:txtBusqueda"] = _texto_busq(filtros)
        data["formBuscador:varAutos"] = "on" if _verdadero(
            filtros.get("autos")) else ""
        r = await self._post(self._url(PAGE_INICIO), data,
                             referer=self._url(PAGE_INICIO))
        if r.status_code not in (301, 302, 303, 307):
            y = r.text
            if "rf-msgs-err" in y:
                raise ErrorValidacion(_msgs(y) or "el buscador rechazo la consulta")
            if "optResultado" in y:
                self._cargar_resultado(y)
                return
            raise ErrorSite(f"la busqueda no navego a resultados (HTTP {r.status_code}); reintenta")
        loc = self._norm(r.headers.get("location", ""))
        r2 = await self._get(loc if loc.startswith("http") else BASE + loc)
        if r2.status_code != 200:
            raise ErrorSite(f"resultados no disponibles (HTTP {r2.status_code})")
        self._cargar_resultado(r2.text)

    def _cargar_resultado(self, y: str) -> None:
        mv = self._vs_de(y)
        if mv:
            self.vs = mv
        mc = re.search(r"De un total de (\d+) resoluciones", y)
        if mc:
            self.corpus = int(mc.group(1))
        mtot = re.search(r"se obtuvieron (\d+) resultados", y)
        if mtot:
            self.total = int(mtot.group(1))
        elif re.search(r"No se encontraron resultados", y):
            self.total = 0
        else:
            raise ErrorSite("la vista de resultados no trajo total ni filas; reintenta la llamada")
        filas = _parsear_filas(y)
        if filas:
            # reemplazo total: el render completo SIEMPRE es la pagina 1 del
            # conjunto activo (así búsquedas angostas no arrastran filas del
            # landing anterior; _paginar hace update para paginas 2..N)
            self.filas = filas
        elif self.total:
            raise ErrorSite(f"el sitio reporto {self.total} resultados pero la vista no trajo filas; reintenta")
        else:
            # busqueda activa sin resultados: la pagina 1 queda VACIA (no se
            # arrastran filas del landing anterior)
            self.filas = {}
        mpag = re.search(r"maxValue: (\d+)", y)
        self.paginas = int(mpag.group(1)) if mpag else _ceil(self.total, TAM_PAGINA)
        self.html_resultado = y
        self.touch()

    async def _paginar(self, pag: int) -> list[dict]:
        data = {
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": "formBuscador:data1",
            "javax.faces.partial.execute": "formBuscador:data1",
            "javax.faces.partial.render": "formBuscador:panealJur formBuscador:optResultado formBuscador:spinner",
            "formBuscador:data1:page": str(pag),
            "formBuscador": "formBuscador",
            "javax.faces.ViewState": self.vs,
        }
        r = await self._post(self._url(PAGE_RESULTADO), data, ajax=True)
        if r.status_code != 200:
            raise ErrorSite(f"la pagina {pag} no se cargo (HTTP {r.status_code})")
        x = r.text
        mv = self._vs_de(x)
        if mv:
            self.vs = mv
        self.html_resultado = x
        mtot = re.search(r"se obtuvieron (\d+) resultados", x)
        if mtot:
            self.total = int(mtot.group(1))
        rows = _parsear_filas(x)
        if not rows:
            raise ErrorSite(f"la pagina {pag} llego sin filas; la sesion pudo expirar: repite la busqueda")
        mnu = re.search(r'name="formBuscador:spinner" size="10" type="text" value="(\d+)"', x)
        if mnu:
            self.paginas = max(self.paginas, int(mnu.group(1)))
        self.filas.update(rows)
        self.touch()
        return _ordenar(rows)

    # ----------------------------------------------------------- filtros

    def _campos_filtro(self, filtros: dict[str, str]) -> dict[str, str]:
        nivel = "1"
        if filtros.get("nivel"):
            if _plain(str(filtros.get("nivel"))) in ALIAS_TODOS:
                raise ErrorSite(
                    "nivel debe ser 'Corte Suprema' o 'Corte Superior'; si no quieres filtrar por nivel, omite el argumento (default = 'Corte Suprema')")
            nivel = _etiqueta(filtros.get("nivel"), NIVEL_CODES, "nivel") or "1"
        distrito = ""
        if filtros.get("distrito") and \
                _plain(str(filtros.get("distrito"))) not in ALIAS_TODOS:
            distrito = _etiqueta(filtros.get("distrito"), DISTRITOS,
                                 "distrito")
            if nivel != "2":
                raise ErrorSite(
                    "el filtro 'distrito' solo aplica con nivel='Corte Superior'; agrega nivel='Corte Superior' o quita el distrito")
        sala = str(filtros.get("sala", "")).strip()
        if _plain(sala) in ALIAS_TODOS:
            sala = ""
        if sala:
            if nivel != "2":
                raise ErrorSite(
                    "el filtro 'sala' solo aplica con nivel='Corte Superior'; agrega nivel='Corte Superior' o quita la sala")
            if not distrito:
                raise ErrorSite(
                    "el filtro 'sala' requiere 'distrito': usa listar_salas(distrito, especialidad) para el valor EXACTO")
        especialidad = "0"
        if filtros.get("especialidad"):
            especialidad = _etiqueta(filtros.get("especialidad"), ESPECIALIDADES,
                                     "especialidad") or "0"
        tipo_recurso = "0"
        if filtros.get("tipo_recurso"):
            tipo_recurso = _etiqueta(filtros.get("tipo_recurso"), TIPOS_RECURSO,
                                     "tipo_recurso") or "0"
        tipo_resolucion = "0"
        if filtros.get("tipo_resolucion"):
            tipo_resolucion = _etiqueta(filtros.get("tipo_resolucion"),
                                        TIPOS_RESOLUCION,
                                        "tipo_resolucion") or "0"
        anio = str(filtros.get("anio", "")).strip()
        if _plain(anio) in ALIAS_TODOS:
            anio = ""
        if anio and anio not in ANIOS:
            raise ErrorSite(f"anio {anio!r} fuera del catalogo del sitio ({ANIOS[-1]}..{ANIOS[0]}; '1982' incluido)")
        return {
            "nivel": nivel,
            "distrito": distrito or "0",
            "especialidad": especialidad,
            "tipo_recurso": tipo_recurso,
            "tipo_resolucion": tipo_resolucion,
            "anio": anio,
            "sala": sala or "0",
            "expediente": _limpia(filtros.get("expediente")),
            "pretension": _limpia(filtros.get("pretension")),
            "palabra_clave": _limpia(filtros.get("palabra_clave")),
            "ncpp": "on" if _verdadero(filtros.get("ncpp")) else "",
            "nlpt": "on" if _verdadero(filtros.get("nlpt")) else "",
        }

    async def _cadena_dependencias(self) -> None:
        campos = dict(self.campos_filtro or {})
        if campos["distrito"] in ("", "0") and campos["sala"] in ("", "0"):
            return
        await self._rf_value("buCorte", campos["nivel"],
                             "formBuscador:buDistrito formBuscador:buSala")
        if campos["distrito"] not in ("", "0"):
            await self._rf_value("buDistrito", campos["distrito"],
                                 "formBuscador:buSala formBuscador:buEspecialidad",
                                 extra={"formBuscador:buCorte": campos["nivel"]})
        if campos["especialidad"] not in ("", "0"):
            await self._rf_value("buEspecialidad", campos["especialidad"],
                                 "formBuscador:buSala",
                                 extra={"formBuscador:buCorte": campos["nivel"],
                                        "formBuscador:buDistrito":
                                            campos["distrito"]})
        if campos["sala"] not in ("", "0"):
            salas = [op for op in _opciones(self.html_rf, "buSala")
                     if op[0] not in ("", "0")]
            if not salas:
                x = await self._rf_plain("buEspecialidad", campos["especialidad"],
                                         "formBuscador:buSala",
                                         extra={"formBuscador:buCorte":
                                                campos["nivel"],
                                                "formBuscador:buDistrito":
                                                    campos["distrito"]})
                self.html_rf = x
                salas = [op for op in _opciones(x, "buSala")
                         if op[0] not in ("", "0")]
            v = _match_combo(salas, campos["sala"])
            if v is None:
                ej = "; ".join(lab for _c, lab in salas[:6] if lab)
                raise ErrorSite(
                    f"la sala {campos['sala']!r} no esta en el catalogo del distrito/especialidad elegidos; ejemplos: {ej or '(sin salas)'}; lista completa con listar_salas(distrito, especialidad)")
            campos["sala"] = v
        self.campos_filtro = campos

    async def _rf_value(self, source: str, val: str, render: str,
                        extra: dict | None = None) -> None:
        self.html_rf = await self._rf_plain(source, val, render, extra)

    async def _rf_plain(self, source: str, val: str, render: str,
                        extra: dict | None = None) -> str:
        data = {
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": "formBuscador:" + source,
            "javax.faces.partial.execute": "formBuscador:" + source,
            "javax.faces.partial.render": render,
            "formBuscador:" + source: val,
            "javax.faces.behavior.event": "valueChange",
            "org.richfaces.ajax.component": "formBuscador:" + source,
            "formBuscador": "formBuscador",
            "javax.faces.ViewState": self.vs,
        }
        if extra:
            data.update(extra)
        r = await self._post(self._url(PAGE_RESULTADO), data, ajax=True)
        if r.status_code != 200:
            raise ErrorSite(f"la cadena de combos no respondio (HTTP {r.status_code})")
        x = r.text
        mv = self._vs_de(x)
        if mv:
            self.vs = mv
        if "rf-msgs-err" in x:
            det = _msgs(x)
            if "Validation Error" in det:
                det = "un valor de la cadena de combos no es valido; usa los valores EXACTOS de listar_salas (sala) o listar_filtros"
            raise ErrorValidacion(det or "el combo rechazo el valor")
        return x

    async def _aplicar_filtros(self, filtros: dict[str, str]) -> None:
        campos = dict(self.campos_filtro or self._campos_filtro(filtros))
        y = self.html_resultado or ""
        pares_filt = None
        for pares in self._jsfcljs(y, "formBuscador"):
            dic = dict(pares)
            if "forward" in dic or "busqueda" in dic:
                continue
            if len(pares) >= 4 and any(v in CODIGOS_ORDEN for _k, v in pares[1:]):
                pares_filt = pares
                break
        if pares_filt is None:
            raise ErrorSite("el botón de filtros no esta disponible; el sitio puede haber cambiado")
        data: dict[str, str] = {"formBuscador": "formBuscador",
                                "javax.faces.ViewState": self.vs}
        for k, v in pares_filt:
            data[k] = v
        orden = _orden_code(filtros.get("orden"))
        forma = _forma_valor(filtros.get("forma"))
        keys_orden = [k for k, v in pares_filt[1:] if v in CODIGOS_ORDEN]
        keys_forma = [k for k, v in pares_filt[1:] if v in ("ASC", "DESC")]
        if keys_orden:
            data[keys_orden[0]] = orden or "21"
        if keys_forma:
            data[keys_forma[0]] = forma
        autos = "on" if _verdadero(filtros.get("autos")) else ""
        data.update({
            "formBuscador:txtBusqueda": _limpia(str(filtros.get("texto", ""))),
            "formBuscador:buCorte": campos["nivel"],
            "formBuscador:buDistrito": campos["distrito"],
            "formBuscador:buEspecialidad": campos["especialidad"],
            "formBuscador:buNcpp": campos["ncpp"],
            "formBuscador:buNlpt": campos["nlpt"],
            "formBuscador:buPretensionValue": campos["pretension"],
            "formBuscador:buPretensionInput": campos["pretension"],
            "formBuscador:buPalabraClaveValue": campos["palabra_clave"],
            "formBuscador:buPalabraClaveInput": campos["palabra_clave"],
            "formBuscador:buNroExpediente": campos["expediente"],
            "formBuscador:buSala": campos["sala"],
            "formBuscador:buPretensionDelitoSupValue": "",
            "formBuscador:buPretensionDelitoSupInput": "",
            "formBuscador:buTipoRecurso": campos["tipo_recurso"],
            "formBuscador:buTipoResolucion": campos["tipo_resolucion"],
            "formBuscador:buTipoResolucionInput": TIPOS_RESOLUCION.get(
                campos["tipo_resolucion"], "-- Todos --"),
            "formBuscador:buAnio": campos["anio"],
            "formBuscador:varAutos": autos,
            "formBuscador:varAutos2": autos,
            "formBuscador:buOrden": (orden or "21") if keys_orden else "21",
            "formBuscador:buOrdenForma": forma,
            "formBuscador:spinner": "1",
        })
        r = await self._post(self._url(PAGE_RESULTADO), data)
        if r.status_code != 200:
            raise ErrorSite(f"los filtros no se aplicaron (HTTP {r.status_code})")
        yy = r.text
        if "rf-msgs-err" in yy:
            raise ErrorValidacion(_msgs(yy) or "los filtros fueron rechazados; revisa los valores con listar_filtros")
        self._cargar_resultado(yy)

    # ------------------------------------------------------------- laudos

    async def _iniciar_laudos(self) -> None:
        r = await self._get(PAGE_COMERCIALES)
        if r.status_code != 200:
            raise ErrorSite(f"comerciales.xhtml no disponible (HTTP {r.status_code})")
        xc = r.text
        mv = self._vs_de(xc)
        if not mv:
            raise ErrorSite("comerciales sin ViewState; revisa el sitio")
        self.vs = mv
        self.html_comercial = xc
        self.pares_laudos = []
        for pares in self._jsfcljs(xc, "formTributario"):
            dic = dict(pares)
            if dic.get("forward") == "buscar" and len(pares) >= 3:
                self.pares_laudos = pares
                break
        if not self.pares_laudos:
            raise ErrorSite("el botón de busqueda de laudos no esta disponible")

    async def _buscar_laudos(self, filtros: dict[str, str]) -> None:
        data: dict[str, str] = {"formTributario": "formTributario",
                                "javax.faces.ViewState": self.vs}
        orden = _orden_code(filtros.get("orden"))
        forma = _forma_valor(filtros.get("forma"))
        for k, v in self.pares_laudos:
            data[k] = v
        for k, v in list(data.items()):
            if v in CODIGOS_ORDEN:
                data[k] = orden or "21"
            elif v in ("ASC", "DESC"):
                data[k] = forma
        anio = str(filtros.get("anio", "")).strip()
        if _plain(anio) in ALIAS_TODOS:
            anio = ""
        if anio and anio not in ANIOS:
            raise ErrorSite(f"anio {anio!r} fuera del catalogo del sitio")
        data.update({
            "formTributario:txtBusqueda": _limpia(str(filtros.get("texto", ""))),
            "formTributario:buAnio": anio,
            "formTributario:buNroExpediente": _limpia(str(filtros.get("expediente", ""))),
            "formTributario:buPalabraClaveValue": _limpia(str(filtros.get("palabra_clave", ""))),
            "formTributario:buPalabraClaveInput": _limpia(str(filtros.get("palabra_clave", ""))),
        })
        r = await self._post(self._url(PAGE_COMERCIALES), data,
                             referer=self._url(PAGE_COMERCIALES))
        if r.status_code not in (301, 302, 303, 307):
            y = r.text
            if "rf-msgs-err" in y:
                raise ErrorValidacion(_msgs(y) or "los filtros de laudos fueron rechazados")
            raise ErrorSite(f"la busqueda de laudos no navego (HTTP {r.status_code})")
        loc = self._norm(r.headers.get("location", ""))
        r2 = await self._get(loc if loc.startswith("http") else BASE + loc)
        if r2.status_code != 200:
            raise ErrorSite(f"resultados de laudos no disponibles (HTTP {r2.status_code})")
        self._cargar_resultado(r2.text)

    # ------------------------------------------------------------- fichas

    async def fila_de(self, index: int) -> dict:
        index = int(index)
        if index < 0:
            raise ErrorSite("index no puede ser negativo")
        if self.total and index >= self.total:
            raise ErrorRango(f"index {index} fuera de rango; el total de la busqueda es " + str(self.total))
        if index in self.filas:
            return self.filas[index]
        pag = index // max(1, TAM_PAGINA) + 1
        if self.total and pag > self.paginas:
            raise ErrorRango(f"pagina {pag} fuera de rango: la busqueda tiene {self.paginas} paginas")
        await self._paginar(pag)
        if index not in self.filas:
            raise ErrorSite(f"el index {index} no quedo en la pagina {pag}; la sesion pudo expirar: repite la busqueda")
        return self.filas[index]

    async def filas_batch(self, indices: list[int]) -> tuple[list[dict], list[dict]]:
        fichas: list[dict] = []
        faltas: list[dict] = []
        por_pag: dict[int, list[int]] = {}
        for i in indices:
            if i < 0 or (self.total and i >= self.total):
                faltas.append({"index": i,
                               "motivo": f"fuera de rango (total {self.total})"})
                continue
            if i not in self.filas:
                por_pag.setdefault(i // max(1, TAM_PAGINA), []).append(i)
        for pag in sorted(por_pag):
            try:
                await self._paginar(pag)
            except ErrorSite as e:
                for i in por_pag[pag]:
                    if i not in self.filas:
                        faltas.append({"index": i, "motivo": str(e)})
        for i in indices:
            if i in self.filas:
                fichas.append(self.filas[i])
        return fichas, faltas


# ---------------------------------------------------------- sesiones (LRU)

_SESIONES: dict[str, Sesion] = {}


def _purgar() -> None:
    ahora = time.time()
    for k in list(_SESIONES):
        if ahora - _SESIONES[k].ts > SESION_TTL:
            s = _SESIONES.pop(k)
            asyncio.ensure_future(s.close())


async def obtener_sesion(ses: str | None) -> Sesion:
    _purgar()
    key = (ses or "").strip()
    s = _SESIONES.get(key) if key else None
    if s is not None:
        s.touch()
        _SESIONES.pop(key)
        _SESIONES[key] = s
        return s
    s = Sesion()
    _SESIONES[s.key] = s
    while len(_SESIONES) > SESION_MAX:
        vieja = _SESIONES.pop(next(iter(_SESIONES)))
        asyncio.ensure_future(vieja.close())
    return s


def _defensa_sesion(ses: str | None) -> dict | None:
    key = (ses or "").strip()
    if key and key not in _SESIONES:
        return {"ok": False, "error": "sesion expirada o no encontrada",
                "pista": "repite la busqueda para obtener un nuevo 'ses'; las sesiones viven 15 min desde su ultima actividad"}
    return None


def _sesion_tiene_busqueda(ses: str | None) -> bool:
    """la sesion ya tiene una busqueda (continuacion sin filtros OK)."""
    key = (ses or "").strip()
    s = _SESIONES.get(key)
    return bool(s and s.modo and (s.filtros or s.total or s.filas))


# ------------------------------------------------------------ autocompletes

class CatalogoSesion:
    """sesion ambient para los autocompletes (carga perezosa)."""

    def __init__(self) -> None:
        self.s: Sesion | None = None
        self.lock = asyncio.Lock()
        self.creada = 0.0


_CAT = CatalogoSesion()


async def _ses_catalogo_locked() -> Sesion:
    """carga/reusa la sesion ambient (SIEMPRE llamar con _CAT.lock tomado)."""
    if _CAT.s is None:
        _CAT.s = Sesion()
    if not _CAT.s.vs or time.time() - _CAT.creada > 600:
        try:
            r = await _CAT.s._get(PAGE_RESULTADO)
            if r.status_code == 200 and _CAT.s._vs_de(r.text):
                _CAT.s.vs = _CAT.s._vs_de(r.text)
                _CAT.s.html_resultado = r.text
                _CAT.creada = time.time()
                return _CAT.s
        except ErrorSite:
            pass
        await _CAT.s._iniciar_principal({"texto": "prescripcion"})
        _CAT.creada = time.time()
    else:
        _CAT.s.touch()
    return _CAT.s


async def _autocomplete(campo: str, q: str) -> list[str]:
    # TODO el flujo va bajo el lock: dos autocompletes en paralelo no pueden
    # pisarse el ViewState de la sesion ambient (verificado con R8)
    async with _CAT.lock:
        c = await _ses_catalogo_locked()
        data = {
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": "formBuscador:" + campo,
            "javax.faces.partial.execute": "formBuscador:" + campo,
            "javax.faces.partial.render": "formBuscador:" + campo,
            "formBuscador:" + campo + ".ajax": "1",
            "formBuscador:" + campo: "formBuscador:" + campo,
            "org.richfaces.ajax.component": "formBuscador:" + campo,
            "formBuscador:" + campo + "Value": q[:100],
            "formBuscador:" + campo + "Input": q[:100],
            "formBuscador": "formBuscador",
            "javax.faces.ViewState": c.vs,
            "formBuscador:txtBusqueda": "",
        }
        r = await c._post(c._url(PAGE_RESULTADO), data, ajax=True,
                          referer=PAGE_RESULTADO)
        if r.status_code == 200:
            mv = c._vs_de(r.text)
            if mv:
                c.vs = mv
            items = re.findall(r'<div class="rf-au-itm[^>]*>([^<]+)', r.text)
            if items:
                return [_limpia(i) for i in items][:50]
        # 0 items o respuesta rara: revalidar la vista ambient y reintentar
        await c._iniciar_principal({"texto": "prescripcion"})
        data["javax.faces.ViewState"] = c.vs
        r = await c._post(c._url(PAGE_RESULTADO), data, ajax=True,
                          referer=PAGE_RESULTADO)
        items = re.findall(r'<div class="rf-au-itm[^>]*>([^<]+)', r.text)
        return [_limpia(i) for i in items][:50]


# --------------------------------------------------------------------- tools

def _envolver(e: Exception, t0: float | None = None) -> dict:
    msg = str(e)
    out: dict[str, Any] = {"ok": False, "error": msg}
    if t0 is not None:
        out["ms"] = _ms_desde(t0)
    if "sesion expirada" in msg:
        out["pista"] = "repite la busqueda para obtener un nuevo 'ses'"
    elif "fuera de rango" in msg or "no puede ser negativo" in msg:
        out["pista"] = "el 'index' es GLOBAL (0-based): revisa 'total' y 'paginas' de la ultima busqueda"
    elif "no es valido" in msg or "del catalogo" in msg or "catalogo" in msg:
        out["pista"] = "obtiene los valores exactos con listar_filtros, buscar_pretensiones, buscar_palabras_clave o listar_salas"
    elif "criterio" in msg.lower():
        out["pista"] = "pasa 'texto' o al menos un filtro (distrito, especialidad, expediente, pretension, palabra_clave, tipo_recurso, anio)"
    elif "HTTP 5" in msg or "no respondio" in msg or "no se cargo" in msg or "satur" in msg:
        out["pista"] = "el servidor del PJ es lento/intermitente: reintenta la llamada en unos segundos"
    elif "no esta disponible" in msg or "puede haber cambiado" in msg:
        out["pista"] = "el portal pudo cambiar de estructura: reporta el caso con el error exacto"
    elif "solo aplica" in msg or "requiere" in msg:
        out["pista"] = msg.split(";")[0][:160]
    else:
        out["pista"] = "reintenta la llamada; si persiste, revisa los filtros con listar_filtros"
    return out


NOTA_SIN_RESULTADOS = "sin resultados; revisa el texto y los filtros (listar_filtros, buscar_pretensiones y buscar_palabras_clave son tus aliados)"


def _nota(cap: int, recort: int, unidad: str = "paginas") -> str:
    return f"se recortaron {recort} {unidad} por el cap de {cap}; pide el resto en otra llamada"


@mcp.tool(title="Buscando resoluciones")
async def buscar_resoluciones(texto: str | None = None,
                              nivel: str | None = None,
                              distrito: str | None = None,
                              especialidad: str | None = None,
                              sala: str | None = None,
                              pretension: str | None = None,
                              palabra_clave: str | None = None,
                              expediente: str | None = None,
                              tipo_recurso: str | None = None,
                              tipo_resolucion: str | None = None,
                              anio: str | None = None,
                              ncpp: str | None = None,
                              nlpt: str | None = None,
                              autos: str | None = None,
                              orden: str | None = None,
                              forma: str | None = None,
                              pagina: int = 1, paginas: int = 1,
                              ses: str | None = None) -> dict:
    """Busca RESOLUCIONES en la Jurisprudencia Nacional Sistematizada del PJ (busca en el CONTENIDO de casaciones, apelaciones, ejecutorias supremas, autos, sentencias de vista, laudos...).

    Args:
        texto: texto libre del contenido de las resoluciones (ej 'prescripcion adquisitiva'); el sitio filtra en AND implicito (2+ palabras = interseccion). Con SOLO texto reproduce la busqueda del home (conjunto completo, ~mas resultados); al agregar filtros el sitio refina su busqueda (como en su UI).
        nivel: 'Corte Suprema' (default) o 'Corte Superior'.
        distrito: distrito judicial (solo con nivel='Corte Superior'; catalogo en listar_filtros categoria='distritos'; ej 'Lima Norte').
        especialidad: especialidad EXACTA (ciltra por nombre completo: 'Penal', 'Familia Penal', 'Laboral', 'Contencioso Adm. Laboral'...; catalogo en listar_filtros categoria='especialidades'; en Corte Superior depende del distrito).
        sala: organo jurisdiccional (sala) EXACTO del distrito (valor de listar_salas(distrito, especialidad); solo Corte Superior; requiere distrito).
        pretension: pretension/delito EXACTA (hallala con buscar_pretensiones; ej 'Prescripcion Adquisitiva de Dominio'). Solo Corte Suprema.
        palabra_clave: palabra clave EXACTA (busca con buscar_palabras_clave; ej 'CONFIRMARON').
        expediente: numero del expediente (ej '00258-2024' o '00258-2024-0-0905-JR-CI-02').
        tipo_recurso: tipo de recurso EXACTO (ej 'Casacion', 'Nulidad'; catalogo en listar_filtros categoria='tipos_recurso').
        tipo_resolucion: 'Auto', 'Ejecutoria Suprema' o 'Sentencia de Vista' (exacto).
        anio: anio de la resolucion ('1998'..'2026' o '1982').
        ncpp: 'si' para solo resoluciones del Nuevo Codigo Procesal Penal (con especialidad='Penal').
        nlpt: 'si' para solo resoluciones de la Nueva Ley Procesal del Trabajo (con especialidad='Laboral').
        autos: 'si' para incluir autos calificatorios.
        orden: 'Nro de Recurso', 'Anio', 'Pretension/Delito', 'Tipo de Resolucion', 'Sala Suprema' o 'Fecha Resolucion' (default; el sitio ordena por fecha DESC).
        forma: 'ASC' o 'DESC' (default).
        pagina: numero de pagina (1-based; 'index' es GLOBAL y continuo entre paginas; el sitio sirve 10 filas por pagina).
        paginas: paginas consecutivas a traer en UNA llamada (1..6; ~8-12 s cada una; si pides mas de 6 hay 'nota' de recorte).
        ses: sesion de una busqueda previa para la pagina siguiente o para re-buscar (filtros distintos re-ejecutan conservando el 'ses'; un 'ses' desconocido con filtros lanza busqueda NUEVA con id nuevo y lo declara en 'nota').

    Devuelve:
        {'ok', 'ses', 'modo', 'total', 'paginas', 'pagina', 'tamano_pagina' (10; dato del sitio), 'siguiente_pagina', 'filas': [{index, recurso, expediente, tipo_resolucion, fecha_resolucion, sumilla (COMPLETA, sin truncar), palabras_clave, especialidad?, organo_jurisdiccional? (en filas de instancia), sala_suprema? (en filas de Corte Suprema), pretension_delito, norma_derecho_interno?, uuid, url_pdf}]}. El sitio omite 'especialidad/organo' en filas de Corte Suprema y publica 'sala_suprema/norma_derecho_interno' en su lugar (y viceversa): son datos alternativos del mismo campo. Los 'index' se usan con ficha_resolucion.
    """
    t0 = time.perf_counter()
    filtros: dict[str, str] = {}
    for k, v in (("texto", texto), ("nivel", nivel), ("distrito", distrito),
                 ("especialidad", especialidad), ("sala", sala),
                 ("pretension", pretension), ("palabra_clave", palabra_clave),
                 ("expediente", expediente), ("tipo_recurso", tipo_recurso),
                 ("tipo_resolucion", tipo_resolucion), ("anio", anio),
                 ("ncpp", ncpp), ("nlpt", nlpt), ("orden", orden),
                 ("forma", forma)):
        if v is not None and str(v).strip():
            filtros[k] = str(v).strip()
    if _verdadero(autos):
        filtros["autos"] = "si"
    criterios = set(filtros) - {"orden", "forma", "autos"}
    ses_pedida = (ses or "").strip()
    ses_muerta = bool(ses_pedida) and ses_pedida not in _SESIONES
    if not criterios:
        tiene_busqueda = not ses_muerta and _sesion_tiene_busqueda(ses)
        if not tiene_busqueda:
            guardia = _defensa_sesion(ses)
            if guardia:
                return {**guardia, "ms": _ms_desde(t0)}
            return {"ok": False,
                    "error": "ingresa al menos un criterio de busqueda",
                    "pista": "ej texto='prescripcion adquisitiva' o expediente='00258-2024-0-0905-JR-CI-02'; los filtros exactos estan en listar_filtros",
                    "ms": _ms_desde(t0)}
    guardia = _defensa_sesion(ses)
    if guardia and max(1, pagina) > 1:
        return {**guardia, "ms": _ms_desde(t0)}
    try:
        s = await obtener_sesion(ses)
        if not criterios and s.filtros is not None:
            filtros = {}
        paginas_real = max(1, min(paginas, MAX_PAGINAS))
        recort = paginas - paginas_real if paginas > MAX_PAGINAS else 0
        filas = await con_watchdog(
            s.buscar("principal", filtros, max(1, pagina), paginas_real),
            "buscar_resoluciones")
        for f in filas:
            if f.get("uuid"):
                f["url_pdf"] = SERVE_PDF + "?uuid=" + f["uuid"]
        out = _cfg(s, "principal", pagina, paginas_real, len(filas))
        if recort > 0:
            out["nota"] = _nota(MAX_PAGINAS, recort)
        if s.aviso_paginas:
            out["nota"] = (out.get("nota", "") + "; " + s.aviso_paginas).strip("; ")
        if ses_muerta:
            nota = f"la sesion {ses_pedida!r} no estaba activa (vencida): buscar_resoluciones ejecuto una busqueda NUEVA con el id {s.key!r}"
            out["nota"] = (out.get("nota", "") + "; " + nota).strip("; ")
        out["filas"] = filas
        if s.total == 0:
            out["nota"] = (out.get("nota", "") + "; " + NOTA_SIN_RESULTADOS
                           ).strip("; ") if out.get("nota") else NOTA_SIN_RESULTADOS
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


def _cfg(s: Sesion, modo: str, pagina: int, paginas: int, nf: int) -> dict:
    out: dict[str, Any] = {
        "ok": True, "ses": s.key, "modo": modo, "total": s.total,
        "paginas": s.paginas, "pagina": pagina, "filas": nf,
        "tamano_pagina": TAM_PAGINA,
    }
    if nf and (pagina - 1) * TAM_PAGINA + nf < s.total:
        out["siguiente_pagina"] = pagina + paginas
    return out


@mcp.tool(title="Buscando laudos")
async def buscar_laudos(texto: str | None = None,
                        anio: str | None = None,
                        expediente: str | None = None,
                        palabra_clave: str | None = None,
                        orden: str | None = None,
                        forma: str | None = None,
                        pagina: int = 1, paginas: int = 1,
                        ses: str | None = None) -> dict:
    """Busca resoluciones de ANULACION y EJECUCION DE LAUDOS ARBITRALES (buscador 'Comerciales' del home del PJ). Usa la misma vista de resultados; la sesion es del buscador de laudos.

    Args:
        texto: texto del contenido de las resoluciones de laudos (ej 'anulacion del laudo').
        anio: anio de la resolucion ('2026'..'1998' o '1982'; catalogo propio del buscador de laudos).
        expediente: numero del expediente.
        palabra_clave: palabra clave EXACTA (lista con buscar_palabras_clave).
        orden: 'Nro de Recurso', 'Anio', 'Pretension/Delito', 'Tipo de Resolucion', 'Sala Suprema' o 'Fecha Resolucion' (default; el buscador de laudos ordena por fecha DESC).
        forma: 'ASC' o 'DESC' (default).
        pagina: numero de pagina (1-based; 10 filas por pagina).
        paginas: paginas consecutivas (1..6).
        ses: sesion de una busqueda previa de laudos para la pagina siguiente.

    Devuelve:
        {'ok', 'ses', 'modo': 'laudos', 'total', 'paginas', 'pagina', 'tamano_pagina', 'filas': [mismas filas de buscar_resoluciones; sumilla COMPLETA]}.
    """
    t0 = time.perf_counter()
    filtros: dict[str, str] = {}
    for k, v in (("texto", texto), ("anio", anio), ("expediente", expediente),
                 ("palabra_clave", palabra_clave), ("orden", orden),
                 ("forma", forma)):
        if v is not None and str(v).strip():
            filtros[k] = str(v).strip()
    criterios = set(filtros) - {"orden", "forma"}
    ses_pedida = (ses or "").strip()
    ses_muerta = bool(ses_pedida) and ses_pedida not in _SESIONES
    if not criterios:
        tiene_busqueda = not ses_muerta and _sesion_tiene_busqueda(ses)
        if not tiene_busqueda:
            guardia = _defensa_sesion(ses)
            if guardia:
                return {**guardia, "ms": _ms_desde(t0)}
            return {"ok": False,
                    "error": "ingresa al menos un criterio para laudos",
                    "pista": "ej texto='anulacion del laudo' o expediente='000001-2012'",
                    "ms": _ms_desde(t0)}
    guardia = _defensa_sesion(ses)
    if guardia and max(1, pagina) > 1:
        return {**guardia, "ms": _ms_desde(t0)}
    try:
        s = await obtener_sesion(ses)
        if not criterios and s.filtros is not None:
            filtros = {}
        paginas_real = max(1, min(paginas, MAX_PAGINAS))
        recort = paginas - paginas_real if paginas > MAX_PAGINAS else 0
        filas = await con_watchdog(
            s.buscar("laudos", filtros, max(1, pagina), paginas_real),
            "buscar_laudos")
        for f in filas:
            if f.get("uuid"):
                f["url_pdf"] = SERVE_PDF + "?uuid=" + f["uuid"]
        out = _cfg(s, "laudos", pagina, paginas_real, len(filas))
        if recort > 0:
            out["nota"] = _nota(MAX_PAGINAS, recort)
        if s.aviso_paginas:
            out["nota"] = (out.get("nota", "") + "; " + s.aviso_paginas).strip("; ")
        if ses_muerta:
            nota = f"la sesion {ses_pedida!r} no estaba activa (vencida): buscar_laudos ejecuto una busqueda NUEVA con el id {s.key!r}"
            out["nota"] = (out.get("nota", "") + "; " + nota).strip("; ")
        out["filas"] = filas
        if s.total == 0:
            out["nota"] = (out.get("nota", "") + "; sin resultados; acota con 'anio' o 'expediente'").strip("; ")
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Obteniendo ficha de resolución")
async def ficha_resolucion(ses: str, index: int) -> dict:
    """FICHA/RESUMEN completo de una resolucion por 'index': recurso, expediente, especialidad, tipo de resolucion, fecha, organo jurisdiccional, pretension/delito, sumilla COMPLETA (sin truncar), palabras clave y URL del PDF.

    Args:
        ses: sesion devuelta por buscar_resoluciones o buscar_laudos.
        index: 'index' global de la fila (0-based; el de las filas de la busqueda).

    Devuelve:
        {'ok', 'ses', 'index', 'fila': {recurso, expediente, tipo_resolucion, fecha_resolucion, sumilla (COMPLETA), palabras_clave, especialidad?/organo_jurisdiccional? o sala_suprema?/norma_derecho_interno? (alternativos del sitio), pretension_delito, uuid, url_pdf}, 'nota' (si la fila no expone PDF)}: el PDF con url_resolucion.
    """
    t0 = time.perf_counter()
    guardia = _defensa_sesion(ses)
    if guardia:
        return {**guardia, "ms": _ms_desde(t0)}
    try:
        s = await obtener_sesion(ses)
        fila = dict(await con_watchdog(s.fila_de(int(index)), "ficha"))
        fila.pop("index", None)
        out: dict[str, Any] = {"ok": True, "ses": s.key, "index": int(index),
                               "fila": fila}
        if fila.get("uuid"):
            fila["url_pdf"] = SERVE_PDF + "?uuid=" + fila["uuid"]
        else:
            out["nota"] = "la fila no expone PDF en el sitio (sin url)"
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Obteniendo fichas en lote")
async def fichas_resoluciones(ses: str, indices: list[int]) -> dict:
    """Lote de fichas/resumenes de la MISMA busqueda en UNA llamada (cap 50; suma sumillas COMPLETAS).

    Args:
        ses: sesion de la busqueda.
        indices: lista de 'index' globales (ej [0, 1, 2]; duplicados se procesan una sola vez).

    Devuelve:
        {'ok', 'ses', 'fichas': [...], 'faltas': [{index, motivo}], 'nota'}: los indices fallidos van a 'faltas' con su motivo.
    """
    t0 = time.perf_counter()
    if not indices:
        return {"ok": False, "error": "indices vacio",
                "pista": "lista de 'index' de la busqueda (ej [0, 1, 2]); revisa 'total' en la ultima busqueda",
                "ms": _ms_desde(t0)}
    guardia = _defensa_sesion(ses)
    if guardia:
        return {**guardia, "ms": _ms_desde(t0)}
    indices = sorted({int(i) for i in indices})
    cap = 50
    recort = len(indices) - cap
    indices = indices[:cap]
    try:
        s = await obtener_sesion(ses)
        if not s.modo or (not s.filas and not s.total and not s.filtros):
            return {"ok": False,
                    "error": "la sesion no tiene una busqueda con resultados",
                    "pista": "corre buscar_resoluciones (o buscar_laudos) primero y usa sus indices",
                    "ms": _ms_desde(t0)}
        fichas, faltas = await con_watchdog(
            s.filas_batch(indices), "fichas_resoluciones")
        for f in fichas:
            if f.get("uuid"):
                f["url_pdf"] = SERVE_PDF + "?uuid=" + f["uuid"]
        out: dict[str, Any] = {"ok": True, "ses": s.key, "fichas": fichas,
                               "faltas": faltas}
        if recort > 0:
            out["nota"] = _nota(cap, recort, "indices")
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Generando URL del PDF")
async def url_resolucion(uuid: str) -> dict:
    """URL PUBLICA de descarga del PDF de una resolucion (ServletDescarga del PJ; sin sesion ni token).

    Args:
        uuid: 'uuid' de la fila o ficha (formato 8-4-4-4-12 hexadecimal).

    Devuelve:
        {'ok', 'url', 'archivo'}: GET publico del PDF; el servidor MCP no lo descarga. Un uuid inexistente/malformado devuelve 500 del sitio.
    """
    t0 = time.perf_counter()
    uid = (uuid or "").strip()
    if not re.fullmatch(RE_UUID, uid):
        return {"ok": False, "error": f"uuid malformado: {uuid!r}",
                "pista": "usa el 'uuid' exacto (8-4-4-4-12) que sale en las filas de buscar_resoluciones/ficha_resolucion",
                "ms": _ms_desde(t0)}
    out = {"ok": True, "url": SERVE_PDF + "?uuid=" + uid,
           "archivo": "resolucion.pdf",
           "nota": "descarga directa del PDF publico del PJ"}
    out["ms"] = _ms_desde(t0)
    return out


async def _excel_impl(s: Sesion) -> dict:
    """cuerpo del exportar_excel (corre bajo el lock de sesion)."""
    async with s.lock:
        y = s.html_resultado or ""
        opc = s._jsfcljs(y, "formBuscador")
        btn = None
        for pares in opc:
            if len(pares) == 1 and pares[0][0].startswith("formBuscador:"):
                btn = pares[0][0]
                break
        if not btn:
            await s._paginar(1)
            for pares in s._jsfcljs(s.html_resultado, "formBuscador"):
                if len(pares) == 1 and pares[0][0].startswith("formBuscador:"):
                    btn = pares[0][0]
                    break
        if not btn:
            raise ErrorSite("el botón Exportar no esta en la vista; el sitio puede haber cambiado")
        data: dict[str, str] = {"formBuscador": "formBuscador",
                                "javax.faces.ViewState": s.vs, btn: btn}
        r = await s._post(s._url(PAGE_RESULTADO), data,
                          referer=s._url(PAGE_RESULTADO))
        ctype = r.headers.get("content-type", "")
        if r.status_code != 200 or "excel" not in ctype:
            raise ErrorSite(f"el sitio no sirvio el excel (HTTP {r.status_code}, content-type {ctype!r})")
        raw = r.content
        if len(raw) > EXCEL_MAX_BYTES:
            return {"ok": False,
                    "error": f"el XLS pesa {len(raw)} bytes y supera el cap de {EXCEL_MAX_BYTES}",
                    "pista": "reduce la busqueda con filtros (anio, especialidad, expediente) y repite; NADA fue truncado: no se sirve el archivo",
                    "tamano_bytes": len(raw), "cap_bytes": EXCEL_MAX_BYTES}
        b64 = base64.b64encode(raw).decode("ascii")
        return {"ok": True, "ses": s.key, "formato": "xls",
                "tamano_bytes": len(raw), "base64": b64,
                "nota": "XLS de la busqueda (el sitio exporta la totalidad del resultado)"}


@mcp.tool(title="Exportando a Excel")
async def exportar_excel(ses: str) -> dict:
    """Exporta la BUSQUEDA ACTUAL de la sesion a EXCEL (XLS) en base64 (mismo boton 'Exportar' del sitio).

    Args:
        ses: sesion devuelta por buscar_resoluciones o buscar_laudos.

    Devuelve:
        {'ok', 'ses', 'formato': 'xls', 'tamano_bytes', 'base64', 'nota'}: decodifica base64 y guarda con extension .xls. Si supera el cap declarado devuelve ok:false CON el tamano y la pista de reduccion (NUNCA trunca el archivo).
    """
    t0 = time.perf_counter()
    guardia = _defensa_sesion(ses)
    if guardia:
        return {**guardia, "ms": _ms_desde(t0)}
    try:
        s = await obtener_sesion(ses)
        if not s.modo or (not s.filas and not s.total and not s.filtros):
            return {"ok": False,
                    "error": "la sesion no tiene una busqueda con resultados",
                    "pista": "corre buscar_resoluciones (o buscar_laudos) primero",
                    "ms": _ms_desde(t0)}
        out = await con_watchdog(_excel_impl(s), "exportar_excel")
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Buscando pretensiones")
async def buscar_pretensiones(q: str) -> dict:
    """Sugerencias del AUTOCOMPLETE de Pretension/Delito (inventario sistematizado del PJ). Los valores EXACTOS se usan como filtro 'pretension' en buscar_resoluciones (solo Corte Suprema).

    Args:
        q: texto inicial (1+ caracteres; ej 'prescripcion', 'indemnizacion', 'pension').

    Devuelve:
        {'ok', 'q', 'sugerencias': [...], 'nota'}: hasta 50 sugerencias.
    """
    t0 = time.perf_counter()
    q = (q or "").strip()
    if not q:
        return {"ok": False, "error": "q vacio",
                "pista": "texto inicial del autocomplete (ej q='prescripcion')",
                "ms": _ms_desde(t0)}
    if len(q) > 100:
        q = q[:100]
    try:
        sug = await con_watchdog(_autocomplete("buPretension", q),
                                 "buscar_pretensiones")
        out: dict[str, Any] = {"ok": True, "q": q, "sugerencias": sug}
        if not sug:
            out["nota"] = "sin sugerencias; prueba menos caracteres"
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Buscando palabras clave")
async def buscar_palabras_clave(q: str) -> dict:
    """Sugerencias del AUTOCOMPLETE de Palabras Clave (sistematizacion de sumillas del PJ). Los valores EXACTOS se usan como 'palabra_clave' en buscar_resoluciones y buscar_laudos.

    Args:
        q: texto inicial (1+ caracteres; ej 'indemnizacion', 'reivindicacion').

    Devuelve:
        {'ok', 'q', 'sugerencias': [...], 'nota'}: hasta 50 sugerencias.
    """
    t0 = time.perf_counter()
    q = (q or "").strip()
    if not q:
        return {"ok": False, "error": "q vacio",
                "pista": "texto inicial del autocomplete (ej q='indemnizacion')",
                "ms": _ms_desde(t0)}
    if len(q) > 100:
        q = q[:100]
    try:
        sug = await con_watchdog(_autocomplete("buPalabraClave", q),
                                 "buscar_palabras_clave")
        out: dict[str, Any] = {"ok": True, "q": q, "sugerencias": sug}
        if not sug:
            out["nota"] = "sin sugerencias; prueba menos caracteres"
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


# --------------------------------------------------------------- salas/combos

_SALAS_CACHE: dict[tuple[str, str], tuple[float, list[tuple[str, str]]]] = {}
_SALAS_LOCK = asyncio.Lock()


async def _salas_de(distrito: str, especialidad: str) -> list[tuple[str, str]]:
    s = Sesion()
    try:
        await s._iniciar_principal({"texto": "prescripcion"})
        await s._rf_value("buCorte", "2",
                          "formBuscador:buDistrito formBuscador:buSala")
        await s._rf_value("buDistrito", distrito,
                          "formBuscador:buSala formBuscador:buEspecialidad",
                          extra={"formBuscador:buCorte": "2"})
        x = await s._rf_plain("buEspecialidad", especialidad,
                              "formBuscador:buSala",
                              extra={"formBuscador:buCorte": "2",
                                     "formBuscador:buDistrito": distrito})
        return [op for op in _opciones(x, "buSala")
                if op[0] not in ("", "0")]
    finally:
        asyncio.ensure_future(s.close())


@mcp.tool(title="Listando salas")
async def listar_salas(distrito: str, especialidad: str) -> dict:
    """Lista los ORGANOS JURISDICCIONALES (salas) de un distrito y especialidad (corte superior). El 'nombre' es el filtro 'sala' EXACTO con nivel='Corte Superior'.

    Args:
        distrito: distrito judicial (ej 'Lima Norte'; catalogo en listar_filtros categoria='distritos').
        especialidad: especialidad (ej 'Civil'; catalogo en listar_filtros categoria='especialidades').

    Devuelve:
        {'ok', 'distrito', 'especialidad', 'salas': [{codigo, nombre}], 'nota'}: usa el 'nombre' EXACTO como filtro 'sala' (o el 'codigo'); ~15 s la primera vez, 30 min de cache.
    """
    t0 = time.perf_counter()
    d_cod = ""
    e_cod = ""
    try:
        d_cod = _etiqueta(distrito, DISTRITOS, "distrito")
        if not d_cod:
            return {"ok": False, "error": "distrito invalido",
                    "pista": "catalogo en listar_filtros(categoria='distritos')",
                    "ms": _ms_desde(t0)}
        e_cod = _etiqueta(especialidad, ESPECIALIDADES, "especialidad")
        if not e_cod:
            return {"ok": False, "error": "especialidad invalida",
                    "pista": "catalogo en listar_filtros(categoria='especialidades')",
                    "ms": _ms_desde(t0)}
        key = (d_cod, e_cod)
        async with _SALAS_LOCK:
            hit = _SALAS_CACHE.get(key)
            salas = []
            if hit and time.time() - hit[0] < 1800:
                salas = hit[1]
            if not salas:
                salas = await con_watchdog(
                    _salas_de(d_cod, e_cod), "listar_salas")
                _SALAS_CACHE[key] = (time.time(), list(salas))
        cat = [{"codigo": cod, "nombre": nom} for cod, nom in salas]
        out = {"ok": True, "distrito": DISTRITOS[d_cod],
               "especialidad": ESPECIALIDADES[e_cod], "salas": cat,
                "nota": ("usa el 'nombre' EXACTO como filtro 'sala' con nivel='Corte Superior' y el mismo distrito/especialidad" if cat else "el distrito/especialidad no expone salas en el sitio")}
        out["ms"] = _ms_desde(t0)
        return out
    except (ErrorSite, ErrorValidacion) as e:
        return _envolver(e, t0)


@mcp.tool(title="Listando catálogos")
async def listar_filtros(categoria: str = "todos") -> dict:
    """Valores EXACTOS para los filtros del buscador: nivel, distritos, especialidades, tipos_recurso, tipos_resolucion, anios y orden. Los errores de match son EXACTOS (por ejemplo 'Penal' no confunde con 'Familia Penal').

    Args:
        categoria: 'todos' (resumen) o 'nivel', 'distritos', 'especialidades', 'tipos_recurso', 'tipos_resolucion', 'anios', 'orden'.

    Devuelve:
        {'ok', 'categoria', ...}: los valores alimentan los filtros de buscar_resoluciones y buscar_laudos. Nota: en Corte Superior las especialidades y salas disponibles dependen del distrito (listar_salas). 'Todos'/'0' equivale a no filtrar.
    """
    t0 = time.perf_counter()
    cat = _plain(categoria) or "todos"
    known = ("todos", "nivel", "distritos", "especialidades", "tipos_recurso",
             "tipos_resolucion", "anios", "orden")
    if cat not in known:
        return {"ok": False, "error": f"categoria {categoria!r} desconocida",
                "pista": "usa: " + ", ".join(known),
                "ms": _ms_desde(t0)}
    mapeo: dict[str, list[tuple[str, str]]] = {
        "nivel": [("1", "Corte Suprema"), ("2", "Corte Superior")],
        "distritos": sorted(DISTRITOS.items(), key=lambda kv: int(kv[0])),
        "especialidades": list(ESPECIALIDADES.items()),
        "tipos_resolucion": list(TIPOS_RESOLUCION.items()),
        "tipos_recurso": sorted(TIPOS_RECURSO.items(), key=lambda kv: int(kv[0])),
        "anios": [(a, a) for a in ANIOS],
        "orden": list(ORDEN_ETIQUETAS),
    }
    if cat == "todos":
        catalogos = {clave: [{"codigo": c, "nombre": n} for c, n in pares]
                     for clave, pares in mapeo.items()}
        return {"ok": True, "categoria": "todos", "catalogos": catalogos,
                "nota": "en Corte Superior pide especialidades/salas del distrito con listar_salas; 'Todos'/'0' no filtra",
                "ms": _ms_desde(t0)}
    pares = mapeo[cat]
    return {"ok": True, "categoria": cat,
            "valores": [{"codigo": c, "nombre": n} for c, n in pares],
            "ms": _ms_desde(t0)}


# ---------------------------------------------------------- contrato cerrado
# El esquema sin 'additionalProperties' (default true) permite que un cliente
# envie argumentos desconocidos que el SDK descarta en silencio antes de la
# tool. Cerramos el contrato: el esquema publica SOLO los args soportados.
for _t in mcp._tool_manager.list_tools():
    _t.parameters["additionalProperties"] = False


# ---------------------------------------------------------------------- main

def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
