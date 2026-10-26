"""
ORDER SALES NEWSLETTER
Gera e envia a newsletter diária de Ordenado para gerentes e CEO.
Visual: modelo novo_mod_rep.py (masthead MARCOLIN · EYEWEAR BRASIL).

Seções:
  §01 Visão Geral     — KPIs YTD + indicadores do dia
  §02 Por Gerente     — YTD 2026 vs 2025 YTD vs 2025 Total
  §03 Mês × Mês      — semanas do mês atual vs SPLY
  §04 Semana × Semana — dias da semana atual vs SPLY
  §05 Regional        — YTD por gerente × região (distrito SAP)
  §06 Por Marca       — YTD por segmento × marca

Execução:
  python orders_newsletter.py
"""

from paths_config import data_path, out_dir
import os
import warnings
from collections import defaultdict
from datetime import datetime, date as _date, timedelta

import openpyxl
import win32com.client

warnings.filterwarnings("ignore")
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from database_gerentes_e_reps import (
    load_key_to_manager, load_clients,
    BRAND_SEGMENT, BRAND_NAMES, EXCLUDED_MANAGERS,
    fmt,
)

# ── Referência D-1 ─────────────────────────────────────────────────────────────
# A newsletter sempre reporta os valores fechados do dia anterior (D-1), mesmo
# que a base Ordenado já contenha lançamentos do dia corrente — evita mostrar
# um dia parcial/ainda em fechamento. Isso sobrescreve localmente as constantes
# de database_gerentes_e_reps.py (baseadas em date.today()), sem afetar as
# demais newsletters que também importam esse módulo.
_TODAY    = _date.today() - timedelta(days=1)
YTD_YEAR  = _TODAY.year
YTD_MONTH = _TODAY.month
YTD_DAY   = _TODAY.day
LY_YEAR   = YTD_YEAR - 1

def _ytd_day_ok(cal_date) -> bool:
    """Mesma lógica de database_gerentes_e_reps._ytd_day_ok, mas usando o
    YTD_DAY recalculado para D-1 (ver comentário acima)."""
    if cal_date is None:
        return False
    if isinstance(cal_date, datetime):
        d = cal_date.date()
    elif isinstance(cal_date, _date):
        d = cal_date
    else:
        return False
    return d.day <= YTD_DAY

# ── Configuração ───────────────────────────────────────────────────────────────
RECIPIENTS = [
    "mpacheco@marchon.com.br",
    "gbaldassarre@marcolin.com",
    "lcruz@marcolin.com",
    "rserafim@marchon.com.br",
    "aluis@vsp.com",
    "pvieira@vsp.com"
]

OUTPUT_DIR  = out_dir("consolidated_newsletters")
TITLE_FONT  = "font-family:'Arial Narrow',Arial,sans-serif"
BODY_FONT   = "font-family:Arial,sans-serif"

# Cor dos segmentos (fundo da linha de cabeçalho)
SEG_COLOR = {
    "DIFFUSION": "#1a3a1a",
    "LUXURY":    "#3a1a1a",
    "THELIOS":   "#1a3a5c",
}
SEG_ORDER = ["DIFFUSION", "LUXURY", "THELIOS"]

# Nomes amigáveis dos gerentes
MGR_LABEL = {
    "MB": "MB",
    "MDB": "MB",
}

REGION_NAMES = {
    "SP": "São Paulo",    "RJ": "Rio de Janeiro",   "MG": "Minas Gerais",
    "ES": "Espírito Santo","RS": "Rio Grande do Sul","SC": "Santa Catarina",
    "PR": "Paraná",       "BA": "Bahia",             "CE": "Ceará",
    "PE": "Pernambuco",   "MA": "Maranhão",          "PB": "Paraíba",
    "RN": "Rio Grande do Norte","AL": "Alagoas",     "SE": "Sergipe",
    "PI": "Piauí",        "GO": "Goiás",             "DF": "Distrito Federal",
    "MT": "Mato Grosso",  "MS": "Mato Grosso do Sul","AM": "Amazonas",
    "PA": "Pará",         "AC": "Acre",              "RO": "Rondônia",
    "RR": "Roraima",      "AP": "Amapá",             "TO": "Tocantins",
    "SU": "Sul",          "NE": "Nordeste",          "CO": "Centro-Oeste",
    "NO": "Norte",
}

_ONE_CHAR_REGION = {"N": "NO", "S": "SU", "C": "CO"}

def _extract_uf(dist):
    """Extrai código de estado/região a partir do distrito SAP (prefixo 2 chars fixo).

    Ex: BASP11 → SP | BANE02 → NE | BAN01 → NO (N+dígito = 1-char region)
    """
    if len(dist) >= 4:
        # Quando o 4º char é dígito, o código de região tem só 1 char (ex: BAN01 → N → Norte)
        if dist[3].isdigit():
            one = dist[2].upper()
            return _ONE_CHAR_REGION.get(one, one)
        return dist[2:4].upper()
    return dist.upper()

def load_sp_division():
    """Lê Data/Divisão SP.xlsx e retorna {código_distrito: "SP1"|"SP2"|"INTERIOR"}.

    A coluna C (código) é o mesmo campo usado como Sales District Id em
    Ordenado.xlsx/Ordenado_2025.xlsx (ex.: BASP09, BASP11). Quando um código
    aparece em mais de uma linha com divisões diferentes, prevalece a divisão
    majoritária; empates e códigos sem linha na planilha resolvem para
    "INTERIOR".
    """
    wb = openpyxl.load_workbook(data_path("Divisão SP.xlsx"), read_only=True, data_only=True)
    ws = wb.active
    counts = defaultdict(lambda: defaultdict(int))
    for r in list(ws.iter_rows(values_only=True))[1:]:
        cidade, divisao, codigo = r[0], r[1], r[2]
        if not codigo or not divisao:
            continue
        code = str(codigo).strip().upper()
        div  = str(divisao).strip().upper()
        if div not in ("SP1", "SP2", "INTERIOR"):
            continue
        counts[code][div] += 1

    sp_division = {}
    for code, div_counts in counts.items():
        max_count = max(div_counts.values())
        leaders = {d for d, c in div_counts.items() if c == max_count}
        sp_division[code] = "INTERIOR" if len(leaders) > 1 or "INTERIOR" in leaders else leaders.pop()
    return sp_division

_SP_GROUP_LABEL = {"SP1": ("SP1", "SP1"), "SP2": ("SP2", "SP2")}

def _region_group(dist, sp_division):
    """Retorna (chave, rótulo) do grupo regional resumido de um distrito.

    São Paulo é dividido em SP1, SP2 e Interior conforme `sp_division`
    (ver load_sp_division, fonte: Data/Divisão SP.xlsx); códigos sem
    classificação na planilha caem em Interior. As demais UFs/regiões
    usam o agrupamento de _extract_uf.
    """
    uf = _extract_uf(dist)
    if uf == "SP":
        div = sp_division.get(dist.upper(), "INTERIOR")
        return _SP_GROUP_LABEL.get(div, ("SP_INT", "Interior de São Paulo"))
    return uf, REGION_NAMES.get(uf, uf)

def _mgr(name):
    u = name.upper()
    return MGR_LABEL.get(u, name.title())

# Subregiões de Leandro Marinho (Norte/Centro-Oeste) e Rafael Leiro (Nordeste),
# definidas por UF real do cliente (o distrito SAP não distingue UF dentro de
# NE/CO — por isso usa-se o estado do cliente). Cada gerente só recebe dados
# das UFs de sua territorial, então a mesma tabela combinada serve os dois.
_TERRITORIAL_GROUPS = [
    ("NORTE", "NORTE", {"AM", "PA", "AC", "RO", "RR", "AP"}),
    ("NE1",   "NE1",   {"CE", "PI", "MA"}),
    ("NE2",   "NE2",   {"RN", "PE", "PB"}),
    ("NE3",   "NE3",   {"AL", "SE", "BA"}),
    ("CO1",   "CO1",   {"GO", "DF", "TO"}),
    ("CO2",   "CO2",   {"MT", "MS"}),
]
_UF_TO_TERRITORIAL_GROUP = {uf: (key, label) for key, label, ufs in _TERRITORIAL_GROUPS for uf in ufs}
_TERRITORIAL_GROUP_ORDER = {key: i for i, (key, _, _) in enumerate(_TERRITORIAL_GROUPS)}

def _territorial_region_group(uf):
    """Agrupa a UF real do cliente na subregião de Leandro Marinho/Rafael Leiro
    (fallback: a própria UF)."""
    return _UF_TO_TERRITORIAL_GROUP.get(uf, (uf, REGION_NAMES.get(uf, uf)))

def _bruno_region_group(uf):
    """Bruno Tepedino: performance detalhada por estado (UF real do cliente)."""
    return uf, REGION_NAMES.get(uf, uf)

# Lucas Gomes: ordem fixa SP1/SP2/SP Interior (com subtotal).
# Luciano Pereira: RJ/MG/ES. Cada gerente só recebe dados de sua territorial,
# então a mesma tabela de ordenação combinada serve os dois.
_SUDESTE_SP_KEYS     = ("SP1", "SP2", "SP_INT")
_SUDESTE_GROUP_ORDER = {"SP1": 0, "SP2": 1, "SP_INT": 2, "RJ": 3, "MG": 4, "ES": 5}
_SUDESTE_LABEL_OVERRIDE = {"SP_INT": "SP Interior"}

# ── Helpers ────────────────────────────────────────────────────────────────────

def _iso_week(dt):
    return dt.isocalendar()[1] if dt else None

def _pct_html(a, b, bold=True):
    """Retorna <span> colorido com variação % de a vs b."""
    if not b:
        return '<span style="color:#bbbbbb">&mdash;</span>'
    p   = (a - b) / abs(b) * 100
    c   = "#1a6b3a" if p >= 0 else "#8a1a1a"
    s   = "+" if p >= 0 else ""
    bld = "font-weight:bold;" if bold else ""
    return f'<span style="color:{c};{bld}">{s}{p:.1f}%</span>'

def _pct_html_dark(a, b, bold=True):
    """Retorna <span> colorido com variação % de a vs b, para uso sobre fundos escuros."""
    if not b:
        return '<span style="color:#cccccc">&mdash;</span>'
    p   = (a - b) / abs(b) * 100
    c   = "#7ee787" if p >= 0 else "#ff7b72"
    s   = "+" if p >= 0 else ""
    bld = "font-weight:bold;" if bold else ""
    return f'<span style="color:{c};{bld}">{s}{p:.1f}%</span>'

def _diff_html(a, b):
    """Retorna <span> colorido com diferença numérica a-b."""
    d = a - b
    c = "#1a6b3a" if d >= 0 else "#8a1a1a"
    s = "+" if d >= 0 else ""
    return f'<span style="color:{c};font-weight:bold">{s}{fmt(d)}</span>'

def section_head(num, label):
    return (
        f'<tr><td style="padding:24px 34px 0 34px">'
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        f'<td width="26" style="{TITLE_FONT};font-weight:bold;font-size:12px;color:#bbbbbb;'
        f'padding-bottom:5px;border-bottom:2px solid #111111;vertical-align:bottom">{num}</td>'
        f'<td style="{TITLE_FONT};font-weight:bold;font-size:13px;letter-spacing:2px;text-transform:uppercase;'
        f'color:#111111;padding-bottom:5px;border-bottom:2px solid #111111;vertical-align:bottom">'
        f'&nbsp;&nbsp;{label}</td></tr></table></td></tr>'
    )

# ── Data Loading ───────────────────────────────────────────────────────────────

def _to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, _date):
        return v
    return None

def load_orders(key_to_manager):
    """
    Carrega pedidos de Ordenado.xlsx (2026) e Ordenado_2025.xlsx (2025).
    Filtros alinhados ao Power BI (idêntico a load_orders_by_manager):
      - resp_id (r[1]) não-nulo
      - EXCLUDED_MANAGERS excluídos
      - YTD 2026: month <= YTD_MONTH e _ytd_day_ok() no mês atual
      - YTD 2025: mesmo recorte com _ytd_day_ok()

    Retorna dicionário com:
      ytd26        {mgr: {brand: qty}}  YTD 2026 (alinhado ao PBI)
      ytd25y       {mgr: {brand: qty}}  mesmo período 2025
      ytd25f       {mgr: {brand: qty}}  ano completo 2025
      today_mgr    {mgr: {brand: qty}}  apenas _TODAY
      daily_26     {date: {mgr: qty}}   dias da semana atual 2026
      daily_25     {date: {mgr: qty}}   dias da semana atual 2025 (mesmo nº semana)
      weekly_26    {(year,wk): {mgr: qty}}  semanas do mês atual 2026
      weekly_25    {(year,wk): {mgr: qty}}  semanas do mês atual 2025 (mesmo nº semana)
      mtd_25       {mgr: qty}           mês atual 2025, limitado ao mesmo dia-do-mês de hoje (§03 MTD)
      region_26    {mgr: {district: qty}}   YTD 2026 por região (distrito SAP)
      region_25y   {mgr: {district: qty}}   YTD 2025 por região (distrito SAP)
      region_26_uf {mgr: {UF: qty}}         YTD 2026 por UF real do cliente
      region_25y_uf{mgr: {UF: qty}}         YTD 2025 por UF real do cliente
    """
    ytd26     = defaultdict(lambda: defaultdict(int))
    ytd25y    = defaultdict(lambda: defaultdict(int))
    ytd25f    = defaultdict(lambda: defaultdict(int))
    today_mgr = defaultdict(lambda: defaultdict(int))
    daily_26  = defaultdict(lambda: defaultdict(int))
    daily_25  = defaultdict(lambda: defaultdict(int))
    weekly_26 = defaultdict(lambda: defaultdict(int))
    weekly_25 = defaultdict(lambda: defaultdict(int))
    mtd_25    = defaultdict(int)
    region_26 = defaultdict(lambda: defaultdict(int))
    region_25y= defaultdict(lambda: defaultdict(int))
    region_26_uf = defaultdict(lambda: defaultdict(int))
    region_25y_uf= defaultdict(lambda: defaultdict(int))

    # UF real do cliente (SAP não distingue UF dentro de NE/CO — usado no §05 por gerente)
    cid_to_estado = {cid: info["estado"] for cid, info in load_clients().items()}
    sp_division = load_sp_division()

    today     = _TODAY
    cur_week  = _iso_week(today)
    cur_month = today.month
    excl      = {m.upper() for m in EXCLUDED_MANAGERS}

    # ── 2026 ──────────────────────────────────────────────────────────────────
    print("  Carregando Ordenado.xlsx (2026)...", flush=True)
    wb = openpyxl.load_workbook(data_path("Ordenado.xlsx"), read_only=True)
    for r in wb.active.iter_rows(min_row=3, values_only=True):
        resp_id  = r[1]
        district, brand, cid, qty, month, cal = r[13], r[14], r[10], r[23], r[3], r[18]

        if not resp_id:                              # filtro PBI: Resp Id <> null
            continue
        if not district or not brand or not qty:
            continue
        if not month:
            continue

        m = int(month)
        if m > YTD_MONTH:
            continue

        key = str(district).strip() + str(brand).strip()
        mgr = key_to_manager.get(key)
        if not mgr or mgr.upper() in excl:
            continue

        qty  = int(qty)
        dist = str(district).strip()
        d    = _to_date(cal)
        estado = cid_to_estado.get(str(cid).strip().zfill(10), "") if cid else ""

        # YTD 2026 — alinhado ao PBI (mesmo recorte de load_orders_by_manager)
        if m < YTD_MONTH or (m == YTD_MONTH and _ytd_day_ok(cal)):
            ytd26[mgr][brand] += qty
            region_26[mgr][dist] += qty
            if estado:
                region_26_uf[mgr][estado] += qty

        # Hoje (data exata)
        if d == today:
            today_mgr[mgr][brand] += qty

        # Dia a dia — semana atual (§04 WoW)
        if d and d <= today and d.year == today.year and _iso_week(d) == cur_week:
            daily_26[d][mgr] += qty

        # Semanas do mês atual (§03 MoM)
        if d and d <= today and d.year == today.year and d.month == cur_month:
            weekly_26[(d.year, _iso_week(d))][mgr] += qty

    wb.close()

    # ── 2025 ──────────────────────────────────────────────────────────────────
    print("  Carregando Ordenado_2025.xlsx (2025)...", flush=True)
    wb = openpyxl.load_workbook(data_path("Ordenado_2025.xlsx"), read_only=True)
    for r in wb.active.iter_rows(min_row=3, values_only=True):
        resp_id  = r[1]
        district, brand, cid, qty, month, cal = r[13], r[14], r[10], r[23], r[3], r[18]

        if not resp_id:
            continue
        if not district or not brand or not qty:
            continue

        key = str(district).strip() + str(brand).strip()
        mgr = key_to_manager.get(key)
        if not mgr or mgr.upper() in excl:
            continue

        qty  = int(qty)
        dist = str(district).strip()
        d    = _to_date(cal)
        estado = cid_to_estado.get(str(cid).strip().zfill(10), "") if cid else ""

        # Ano completo 2025
        ytd25f[mgr][brand] += qty

        # YTD 2025 — mesmo recorte _ytd_day_ok (alinhado ao PBI)
        if month:
            m = int(month)
            if m < YTD_MONTH or (m == YTD_MONTH and _ytd_day_ok(cal)):
                ytd25y[mgr][brand] += qty
                region_25y[mgr][dist] += qty
                if estado:
                    region_25y_uf[mgr][estado] += qty

        # Semana × semana — mesma semana ISO em 2025 (§04 WoW)
        if d and d.year == LY_YEAR:
            wk25 = _iso_week(d)
            if wk25 == cur_week:
                daily_25[d][mgr] += qty
            # Mesmo mês (§03 MoM)
            if d.month == cur_month:
                weekly_25[(d.year, wk25)][mgr] += qty
                # MTD — mesmo dia-do-mês de hoje (§03 coluna MTD)
                if d.day <= today.day:
                    mtd_25[mgr] += qty

    wb.close()

    return {
        "ytd26":     dict(ytd26),
        "ytd25y":    dict(ytd25y),
        "ytd25f":    dict(ytd25f),
        "today_mgr": dict(today_mgr),
        "daily_26":  dict(daily_26),
        "daily_25":  dict(daily_25),
        "weekly_26": dict(weekly_26),
        "weekly_25": dict(weekly_25),
        "mtd_25":    dict(mtd_25),
        "region_26": dict(region_26),
        "region_25y":dict(region_25y),
        "region_26_uf": dict(region_26_uf),
        "region_25y_uf":dict(region_25y_uf),
        "sp_division":  sp_division,
    }


# ── HTML Builder ───────────────────────────────────────────────────────────────

def build_html(gen_date, data):
    today     = gen_date
    today_lbl = today.strftime("%d %b %Y").upper()
    today_str = today.strftime("%d/%m/%Y")
    month_lbl = today.strftime("%b %Y").upper()

    ytd26  = data["ytd26"]
    ytd25y = data["ytd25y"]
    ytd25f = data["ytd25f"]
    today_mgr = data["today_mgr"]
    daily_26  = data["daily_26"]
    daily_25  = data["daily_25"]
    weekly_26 = data["weekly_26"]
    weekly_25 = data["weekly_25"]
    mtd_25    = data["mtd_25"]
    region_26 = data["region_26"]
    region_25y= data["region_25y"]
    region_26_uf = data["region_26_uf"]
    region_25y_uf= data["region_25y_uf"]
    sp_division  = data["sp_division"]

    all_mgrs = sorted(
        set(ytd26) | set(ytd25y),
        key=lambda m: -sum(ytd26.get(m, {}).values()),
    )

    # ── Totais gerais ─────────────────────────────────────────────────────────
    def _sum(d):
        return sum(sum(v.values()) for v in d.values())

    total_26  = _sum(ytd26)
    total_25y = _sum(ytd25y)
    total_25f = _sum(ytd25f)
    total_today = sum(sum(v.values()) for v in today_mgr.values())

    # ── Totais por segmento (YTD 2026) ────────────────────────────────────────
    seg_total = {s: 0 for s in SEG_ORDER}
    for mgr_brands in ytd26.values():
        for brand, qty in mgr_brands.items():
            seg = BRAND_SEGMENT.get(brand)
            if seg in seg_total:
                seg_total[seg] += qty

    # ── Hoje por segmento ─────────────────────────────────────────────────────
    seg_today = {s: 0 for s in SEG_ORDER}
    for mgr_brands in today_mgr.values():
        for brand, qty in mgr_brands.items():
            seg = BRAND_SEGMENT.get(brand)
            if seg in seg_today:
                seg_today[seg] += qty

    # ── Semanas do mês atual ──────────────────────────────────────────────────
    cur_month = today.month
    cur_week  = _iso_week(today)

    weeks_26 = sorted(set(weekly_26.keys()))
    weeks_25 = sorted(set(weekly_25.keys()))

    # ── Dias da semana atual ──────────────────────────────────────────────────
    week_days_26 = sorted(daily_26.keys())
    week_days_25 = sorted(daily_25.keys())

    # ── Nomes PT-BR dos dias da semana ────────────────────────────────────────
    DAY_PT = {0:"seg",1:"ter",2:"qua",3:"qui",4:"sex",5:"sáb",6:"dom"}

    # =========================================================================
    # §01 — VISÃO GERAL
    # =========================================================================
    def _seg_share_bar(seg):
        val = seg_total.get(seg, 0)
        pct = val / total_26 * 100 if total_26 else 0
        w   = int(pct / 100 * 80)
        return (
            f'<table width="80" cellpadding="0" cellspacing="0" '
            f'style="background:#e0e0e0;height:8px;display:inline-table;vertical-align:middle">'
            f'<tr><td width="{w}" style="background:{SEG_COLOR[seg]};height:8px"></td><td></td></tr>'
            f'</table>'
            f'<span style="{BODY_FONT};font-size:9px;color:#777;margin-left:4px">{pct:.1f}%</span>'
        )

    seg_rows = ""
    for seg in SEG_ORDER:
        v26  = seg_total.get(seg, 0)
        v25  = sum(sum(b for br, b in mgr_brands.items()
                       if BRAND_SEGMENT.get(br) == seg)
                   for mgr_brands in ytd25y.values())
        seg_rows += (
            f'<tr style="border-bottom:1px solid #e8e8e8">'
            f'<td style="padding:5px 8px;{BODY_FONT};font-size:11px;font-weight:bold;font-style:italic">{seg.title()}</td>'
            f'<td style="padding:4px 8px;vertical-align:middle">{_seg_share_bar(seg)}</td>'
            f'<td style="padding:5px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px">{fmt(v26)}</td>'
            f'<td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(v25)}</td>'
            f'<td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(v26, v25)}</td>'
            f'</tr>'
        )

    # Hoje por gerente (mini-tabela)
    today_rows = ""
    for mgr in all_mgrs:
        mb = today_mgr.get(mgr, {})
        tot = sum(mb.values())
        d26 = sum(v for br, v in mb.items() if BRAND_SEGMENT.get(br) == "DIFFUSION")
        l26 = sum(v for br, v in mb.items() if BRAND_SEGMENT.get(br) == "LUXURY")
        t26 = sum(v for br, v in mb.items() if BRAND_SEGMENT.get(br) == "THELIOS")
        today_rows += (
            f'<tr style="border-bottom:1px solid #eeeeee">'
            f'<td style="padding:4px 8px;{BODY_FONT};font-size:11px;font-weight:600">{_mgr(mgr)}</td>'
            f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#555">{fmt(d26) if d26 else "&mdash;"}</td>'
            f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#555">{fmt(l26) if l26 else "&mdash;"}</td>'
            f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#555">{fmt(t26) if t26 else "&mdash;"}</td>'
            f'<td style="padding:4px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px">{fmt(tot) if tot else "0"}</td>'
            f'</tr>'
        )

    sec01 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="1"
         border="0" style="background:#bbbbbb;border-collapse:separate;margin-bottom:10px">
    <tr>
      <td width="25%" style="background:#111111;padding:12px 14px;text-align:center;vertical-align:middle">
        <div style="{BODY_FONT};font-size:8px;letter-spacing:1.5px;text-transform:uppercase;color:rgba(255,255,255,.7);margin-bottom:3px">Ordenado {YTD_YEAR} YTD</div>
        <div style="{TITLE_FONT};font-weight:bold;font-size:30px;color:#ffffff;line-height:1">{fmt(total_26)}</div>
        <div style="{BODY_FONT};font-size:9px;color:rgba(255,255,255,.6);margin-top:2px">pe&ccedil;as</div>
      </td>
      <td width="25%" style="background:#f5f5f5;padding:12px 14px;text-align:center;vertical-align:middle">
        <div style="{BODY_FONT};font-size:8px;letter-spacing:1.5px;text-transform:uppercase;color:#999;margin-bottom:3px">Ordenado {LY_YEAR} YTD</div>
        <div style="{TITLE_FONT};font-weight:bold;font-size:26px;color:#aaa;line-height:1">{fmt(total_25y)}</div>
        <div style="{BODY_FONT};font-size:9px;color:#bbb;margin-top:2px">refer&ecirc;ncia</div>
      </td>
      <td width="25%" style="background:#ffffff;padding:12px 14px;text-align:center;vertical-align:middle">
        <div style="{BODY_FONT};font-size:8px;letter-spacing:1.5px;text-transform:uppercase;color:#666;margin-bottom:3px">YoY ({YTD_YEAR} vs {LY_YEAR})</div>
        <div style="{TITLE_FONT};font-weight:bold;font-size:26px;line-height:1.1;margin-top:2px">{_pct_html(total_26, total_25y)}</div>
      </td>
      <td width="25%" style="background:#ffffff;padding:12px 14px;text-align:center;vertical-align:middle">
        <div style="{BODY_FONT};font-size:8px;letter-spacing:1px;text-transform:uppercase;color:#666;margin-bottom:3px">Hoje &mdash; {today_str}</div>
        <div style="{TITLE_FONT};font-weight:bold;font-size:30px;color:#333;line-height:1">{fmt(total_today)}</div>
        <div style="{BODY_FONT};font-size:9px;color:#aaa;margin-top:2px">pe&ccedil;as no dia</div>
      </td>
    </tr>
  </table>

  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr valign="top">
      <!-- Segmentos YTD -->
      <td style="padding-right:10px">
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;{BODY_FONT}">
          <thead>
            <tr style="border-top:2px solid #111111">
              <th colspan="5" style="padding:5px 8px;text-align:left;{BODY_FONT};font-size:9px;letter-spacing:1px;text-transform:uppercase;color:#777;font-weight:bold;border-bottom:1px solid #dddddd">YTD por Segmento</th>
            </tr>
            <tr style="border-bottom:1px solid #cccccc;background:#f5f5f5">
              <th style="padding:4px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333">Seg.</th>
              <th style="padding:4px 8px;font-size:8px;font-weight:700;color:#333">Share</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333">{YTD_YEAR} YTD</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888">{LY_YEAR} YTD</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333">YoY</th>
            </tr>
          </thead>
          <tbody>
            {seg_rows}
            <tr style="background:#111111">
              <td colspan="2" style="padding:5px 8px;{BODY_FONT};font-size:11px;font-weight:bold;color:#fff">TOTAL</td>
              <td style="padding:5px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px;color:#fff">{fmt(total_26)}</td>
              <td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(total_25y)}</td>
              <td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(total_26, total_25y)}</td>
            </tr>
          </tbody>
        </table>
      </td>
      <!-- Divisor vertical -->
      <td width="1" style="background:#e0e0e0;padding:0">&nbsp;</td>
      <!-- Hoje por Gerente -->
      <td style="padding-left:10px">
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;{BODY_FONT}">
          <thead>
            <tr style="border-top:2px solid #111111">
              <th colspan="5" style="padding:5px 8px;text-align:left;{BODY_FONT};font-size:9px;letter-spacing:1px;text-transform:uppercase;color:#777;font-weight:bold;border-bottom:1px solid #dddddd">Hoje por Gerente &mdash; {today_str}</th>
            </tr>
            <tr style="border-bottom:1px solid #cccccc;background:#f5f5f5">
              <th style="padding:4px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333">Gerente</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;color:#1a3a1a">Diff.</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;color:#3a1a1a">Lux.</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;color:#1a3a5c">The.</th>
              <th style="padding:4px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333">Total</th>
            </tr>
          </thead>
          <tbody>
            {today_rows}
            <tr style="background:#111111">
              <td style="padding:5px 8px;{BODY_FONT};font-size:11px;font-weight:bold;color:#fff">TOTAL</td>
              <td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#aaa">{fmt(seg_today.get("DIFFUSION",0))}</td>
              <td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#aaa">{fmt(seg_today.get("LUXURY",0))}</td>
              <td style="padding:5px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#aaa">{fmt(seg_today.get("THELIOS",0))}</td>
              <td style="padding:5px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px;color:#fff">{fmt(total_today)}</td>
            </tr>
          </tbody>
        </table>
      </td>
    </tr>
  </table>
</td></tr>"""

    # =========================================================================
    # §02 — POR GERENTE (YTD) com breakdown por segmento
    # =========================================================================
    SEG_SHORT = {"DIFFUSION": "Diff.", "LUXURY": "Lux.", "THELIOS": "The."}

    max_26 = max((sum(v.values()) for v in ytd26.values()), default=1) or 1
    mgr_rows = ""
    for i, mgr in enumerate(all_mgrs):
        v26  = sum(ytd26.get(mgr, {}).values())
        v25y = sum(ytd25y.get(mgr, {}).values())
        v25f = sum(ytd25f.get(mgr, {}).values())
        share = v26 / total_26 * 100 if total_26 else 0
        bar_w = int(v26 / max_26 * 100)
        bg = "#f7f7f5" if i % 2 else "#ffffff"

        # Manager total row
        mgr_rows += (
            f'<tr style="background:{bg};border-bottom:1px solid #e0e0e0">'
            f'<td style="padding:6px 8px;{BODY_FONT};font-size:11.5px;font-weight:700;white-space:nowrap">{_mgr(mgr)}</td>'
            f'<td style="padding:4px 8px;vertical-align:middle;white-space:nowrap">'
            f'<table width="100" cellpadding="0" cellspacing="0" style="background:#e0e0e0;height:10px;display:inline-table;vertical-align:middle">'
            f'<tr><td width="{bar_w}" style="background:#111111;height:10px"></td><td></td></tr></table>'
            f'<span style="{BODY_FONT};font-size:9px;color:#555;margin-left:4px">{share:.1f}%</span></td>'
            f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:14px">{fmt(v26)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(v25y)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(v25f)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(v26, v25y)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(v26, v25f)}</td>'
            f'</tr>'
        )

        # Segment sub-rows: Diff | Lux | The
        for si, seg in enumerate(SEG_ORDER):
            sv26 = sum(v for br, v in ytd26.get(mgr, {}).items() if BRAND_SEGMENT.get(br) == seg)
            sv25 = sum(v for br, v in ytd25y.get(mgr, {}).items() if BRAND_SEGMENT.get(br) == seg)
            seg_c = SEG_COLOR.get(seg, "#333")
            is_last_seg = (si == len(SEG_ORDER) - 1)
            border = "border-bottom:2px solid #cccccc" if is_last_seg else "border-bottom:1px solid #f0f0f0"
            mgr_rows += (
                f'<tr style="background:{bg};{border}">'
                f'<td style="padding:3px 8px 3px 22px;{BODY_FONT};font-size:9.5px;color:{seg_c};font-style:italic">'
                f'&#8627;&nbsp;{SEG_SHORT[seg]}</td>'
                f'<td></td>'
                f'<td style="padding:3px 8px;text-align:right;{TITLE_FONT};font-size:11px;color:#333">'
                f'{fmt(sv26) if sv26 else "&mdash;"}</td>'
                f'<td style="padding:3px 8px;text-align:right;{BODY_FONT};font-size:10px;color:#aaa">'
                f'{fmt(sv25) if sv25 else "&mdash;"}</td>'
                f'<td></td>'
                f'<td style="padding:3px 8px;text-align:right;{BODY_FONT};font-size:10px">'
                f'{_pct_html(sv26, sv25)}</td>'
                f'<td></td>'
                f'</tr>'
            )

    sec02 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">
    <thead><tr style="border-top:2px solid #111111;border-bottom:1px solid #dddddd">
      <th style="padding:5px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Gerente</th>
      <th style="padding:5px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Share %</th>
      <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">{YTD_YEAR} YTD</th>
      <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">{LY_YEAR} SPLY</th>
      <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#aaa;{BODY_FONT}">{LY_YEAR} Total</th>
      <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">YoY (acum.)</th>
      <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">vs Total {LY_YEAR}</th>
    </tr></thead>
    <tbody>
      {mgr_rows}
      <tr style="background:#111111;border-top:2px solid #111111">
        <td colspan="2" style="padding:6px 8px;{BODY_FONT};font-size:11px;font-weight:bold;color:#fff;letter-spacing:1px;text-transform:uppercase">OVERALL</td>
        <td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:14px;color:#fff">{fmt(total_26)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(total_25y)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(total_25f)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(total_26, total_25y)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(total_26, total_25f)}</td>
      </tr>
    </tbody>
  </table>
</td></tr>"""

    # =========================================================================
    # §03 — MÊS × MÊS
    # =========================================================================
    # Alinhar semanas: para cada semana 26 encontrar a equivalente 25
    all_week_keys_26 = sorted(weeks_26)
    # Número de semanas ISO para montar header
    wk_hdrs = ""
    for yk in all_week_keys_26:
        _, wk = yk
        is_cur = (wk == cur_week)
        bg = "#111111" if is_cur else "#eeeeee"
        fg = "#ffffff" if is_cur else "#999999"
        wk_hdrs += (
            f'<td style="padding:6px 8px;text-align:right;background:{bg};color:{fg};'
            f'{BODY_FONT};font-size:9px;font-weight:bold;white-space:nowrap">W{wk:02d}</td>'
        )

    def _mom_row(mgr, label, bold=False):
        cells_26 = cells_25 = cells_var = ""
        tot_26 = 0
        for yk in all_week_keys_26:
            _, wk = yk
            is_cur = (wk == cur_week)
            bg_suf = ";background:#f3f3f1" if is_cur else ""
            v26 = weekly_26.get(yk, {}).get(mgr, 0)
            v25 = weekly_25.get((LY_YEAR, wk), {}).get(mgr, 0)
            tot_26 += v26
            b_style = ";font-weight:bold" if bold else ""
            cells_26 += (f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-size:13px{b_style}{bg_suf}">'
                         f'{fmt(v26) if v26 else "&mdash;"}</td>')
            cells_25 += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888{bg_suf}">'
                         f'{fmt(v25) if v25 else "&mdash;"}</td>')
            cells_var += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:10px{bg_suf}">'
                          f'{_pct_html(v26, v25) if v25 else "&mdash;"}</td>')
        # Total SPLY = mês completo de 2025, incluindo semanas que cruzam a
        # fronteira do mês (ex: 01/jun/2025 cai na semana 22, ausente em jun/2026)
        tot_25_full = sum(v.get(mgr, 0) for v in weekly_25.values())
        tot_25_mtd  = mtd_25.get(mgr, 0)
        return tot_26, tot_25_full, tot_25_mtd, cells_26, cells_25, cells_var

    mom_rows_html = ""
    overall_mom_26 = overall_mom_25 = 0
    overall_cells_26 = overall_cells_25 = overall_cells_var = ("" * len(all_week_keys_26))

    for i, mgr in enumerate(all_mgrs):
        bg = "background:#f7f7f5;" if i % 2 else ""
        tot_26, tot_25, tot_25_mtd, c26, c25, cvar = _mom_row(mgr, mgr)
        mom_rows_html += (
            f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
            f'<td style="padding:6px 8px;{BODY_FONT};font-size:11px;font-weight:600;white-space:nowrap">{_mgr(mgr)}</td>'
            f'{c26}'
            f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px">{fmt(tot_26)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(tot_25_mtd)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(tot_26, tot_25_mtd)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(tot_25)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_diff_html(tot_26, tot_25)}</td>'
            f'</tr>'
            f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
            f'<td style="padding:3px 8px;{BODY_FONT};font-size:9px;color:#888">  SPLY {LY_YEAR}</td>'
            f'{c25}'
            f'<td></td><td></td><td></td><td></td><td></td></tr>'
            f'<tr style="{bg}border-bottom:2px solid #dddddd">'
            f'<td style="padding:3px 8px;{BODY_FONT};font-size:9px;color:#555;font-weight:bold">  Var %</td>'
            f'{cvar}'
            f'<td></td><td></td><td></td><td></td><td></td></tr>'
        )
        overall_mom_26 += tot_26
        overall_mom_25 += tot_25

    # Overall row
    tot_26_all, tot_25_all, tot_25_mtd_all, c26_all, c25_all, cvar_all = _mom_row("__DUMMY__", "__DUMMY__")
    # Recalculate overall across all managers
    cells_26_ov = cells_25_ov = cells_var_ov = ""
    for yk in all_week_keys_26:
        _, wk = yk
        is_cur = (wk == cur_week)
        bg_suf = ";background:#333333" if is_cur else ""
        v26 = sum(weekly_26.get(yk, {}).values())
        yk25 = (LY_YEAR, wk)
        v25 = sum(weekly_25.get(yk25, {}).values())
        cells_26_ov += (f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px;color:#fff{bg_suf}">'
                        f'{fmt(v26) if v26 else "&mdash;"}</td>')
        cells_25_ov += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa{bg_suf}">'
                        f'{fmt(v25) if v25 else "&mdash;"}</td>')

    tot_ov_26 = sum(sum(weekly_26.get(yk, {}).values()) for yk in all_week_keys_26)
    # Mês completo 2025 — soma todas as semanas, inclusive as que cruzam fronteira de mês
    tot_ov_25 = sum(sum(v.values()) for v in weekly_25.values())
    # MTD 2025 — mesmo dia-do-mês de hoje, todos os gerentes
    tot_ov_25_mtd = sum(mtd_25.values())

    sec03 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">
    <thead>
      <tr style="border-top:2px solid #111111;border-bottom:1px solid #dddddd">
        <th style="padding:5px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Gerente</th>
        {wk_hdrs}
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Total {month_lbl}</th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">MTD {LY_YEAR}<br><span style="font-weight:400;font-size:7px;text-transform:none">(at&eacute; {today.strftime('%d/%m')})</span></th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Var % MTD</th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">SPLY {LY_YEAR}</th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Diff MoM</th>
      </tr>
    </thead>
    <tbody>
      {mom_rows_html}
      <tr style="background:#111111;border-top:2px solid #111111">
        <td style="padding:6px 8px;{BODY_FONT};font-size:11px;font-weight:bold;color:#fff;letter-spacing:1px;text-transform:uppercase">OVERALL</td>
        {cells_26_ov}
        <td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:14px;color:#fff">{fmt(tot_ov_26)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(tot_ov_25_mtd)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_pct_html(tot_ov_26, tot_ov_25_mtd)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(tot_ov_25)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_diff_html(tot_ov_26, tot_ov_25)}</td>
      </tr>
    </tbody>
  </table>
</td></tr>"""

    # =========================================================================
    # §04 — SEMANA × SEMANA (dias)
    # =========================================================================
    # Coletar dias de ambos os anos e alinhar por posição na semana (weekday)
    # dias_26: dias com dados em 2026, por weekday 0=seg..6=dom
    days_by_wd_26 = {}   # {weekday: date}
    for d in sorted(daily_26.keys()):
        days_by_wd_26[d.weekday()] = d

    days_by_wd_25 = {}
    for d in sorted(daily_25.keys()):
        days_by_wd_25[d.weekday()] = d

    all_weekdays = sorted(set(days_by_wd_26) | set(days_by_wd_25))

    # Cabeçalho de colunas (dias)
    day_hdrs = ""
    for wd in all_weekdays:
        d26 = days_by_wd_26.get(wd)
        is_today = (d26 == today) if d26 else False
        bg = "#111111" if is_today else "#eeeeee"
        fg = "#ffffff" if is_today else "#999999"
        lbl = f"{DAY_PT.get(wd,'?')}<br><span style='font-size:7px;font-weight:normal'>{d26.strftime('%d/%m') if d26 else '—'}</span>"
        day_hdrs += (f'<td style="padding:6px 8px;text-align:right;background:{bg};color:{fg};'
                     f'{BODY_FONT};font-size:9px;font-weight:bold;white-space:nowrap">{lbl}</td>')

    def _wow_row(mgr, bold=False):
        cells_26 = cells_25 = cells_var = ""
        tot_26 = tot_25 = 0
        for wd in all_weekdays:
            d26 = days_by_wd_26.get(wd)
            d25 = days_by_wd_25.get(wd)
            v26 = daily_26.get(d26, {}).get(mgr, 0) if d26 else 0
            v25 = daily_25.get(d25, {}).get(mgr, 0) if d25 else 0
            is_today = (d26 == today) if d26 else False
            bg_suf = ";background:#f3f3f1" if is_today else ""
            b_style = ";font-weight:bold" if bold else ""
            cells_26 += (f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-size:13px{b_style}{bg_suf}">'
                         f'{fmt(v26) if v26 else ("0" if is_today else "&mdash;")}</td>')
            cells_25 += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888{bg_suf}">'
                         f'{fmt(v25) if v25 else "&mdash;"}</td>')
            cells_var += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:10px{bg_suf}">'
                          f'{_pct_html(v26, v25) if v25 else "&mdash;"}</td>')
            tot_26 += v26
            tot_25 += v25
        return tot_26, tot_25, cells_26, cells_25, cells_var

    wow_rows_html = ""
    for i, mgr in enumerate(all_mgrs):
        bg = "background:#f7f7f5;" if i % 2 else ""
        tot_26, tot_25, c26, c25, cvar = _wow_row(mgr)
        wow_rows_html += (
            f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
            f'<td style="padding:6px 8px;{BODY_FONT};font-size:11px;font-weight:600;white-space:nowrap">{_mgr(mgr)}</td>'
            f'{c26}'
            f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px">{fmt(tot_26)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(tot_25)}</td>'
            f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_diff_html(tot_26, tot_25)}</td>'
            f'</tr>'
            f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
            f'<td style="padding:3px 8px;{BODY_FONT};font-size:9px;color:#888">  SPLY {LY_YEAR}</td>'
            f'{c25}<td></td><td></td><td></td></tr>'
            f'<tr style="{bg}border-bottom:2px solid #dddddd">'
            f'<td style="padding:3px 8px;{BODY_FONT};font-size:9px;color:#555;font-weight:bold">  Var %</td>'
            f'{cvar}<td></td><td></td><td></td></tr>'
        )

    # Overall
    cells_26_ov = cells_25_ov = ""
    tot_ov_26 = tot_ov_25 = 0
    for wd in all_weekdays:
        d26 = days_by_wd_26.get(wd)
        d25 = days_by_wd_25.get(wd)
        v26 = sum(daily_26.get(d26, {}).values()) if d26 else 0
        v25 = sum(daily_25.get(d25, {}).values()) if d25 else 0
        is_today = (d26 == today) if d26 else False
        bg_suf = ";background:#333333" if is_today else ""
        cells_26_ov += (f'<td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:13px;color:#fff{bg_suf}">'
                        f'{fmt(v26) if v26 else "&mdash;"}</td>')
        cells_25_ov += (f'<td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa{bg_suf}">'
                        f'{fmt(v25) if v25 else "&mdash;"}</td>')
        tot_ov_26 += v26
        tot_ov_25 += v25

    sec04 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">
    <thead>
      <tr style="border-top:2px solid #111111;border-bottom:1px solid #dddddd">
        <th style="padding:5px 8px;text-align:left;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Gerente</th>
        {day_hdrs}
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Sem {cur_week:02d} 2026</th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">Sem {cur_week:02d} 2025</th>
        <th style="padding:5px 8px;text-align:right;font-size:8px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">Diff WoW</th>
      </tr>
    </thead>
    <tbody>
      {wow_rows_html}
      <tr style="background:#111111;border-top:2px solid #111111">
        <td style="padding:6px 8px;{BODY_FONT};font-size:11px;font-weight:bold;color:#fff;letter-spacing:1px;text-transform:uppercase">OVERALL</td>
        {cells_26_ov}
        <td style="padding:6px 8px;text-align:right;{TITLE_FONT};font-weight:bold;font-size:14px;color:#fff">{fmt(tot_ov_26)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#aaa">{fmt(tot_ov_25)}</td>
        <td style="padding:6px 8px;text-align:right;{BODY_FONT};font-size:11px">{_diff_html(tot_ov_26, tot_ov_25)}</td>
      </tr>
    </tbody>
  </table>
</td></tr>"""

    # =========================================================================
    # §05 — REGIONAL (por gerente)
    # =========================================================================
    reg_blocks = ""
    for mgr in all_mgrs:
        mgr_u = mgr.upper()
        # Bruno Tepedino, Leandro Marinho e Rafael Leiro usam a UF real do
        # cliente (Clientes.xlsx), pois o distrito SAP não distingue estados
        # dentro de NE/CO.
        if mgr_u in ("BRUNO TEPEDINO", "LEANDRO MARINHO", "RAFAEL LEIRO"):
            r26 = region_26_uf.get(mgr, {})
            r25 = region_25y_uf.get(mgr, {})
        else:
            r26 = region_26.get(mgr, {})
            r25 = region_25y.get(mgr, {})
        all_regs = sorted(set(r26) | set(r25))
        if not all_regs:
            continue
        mgr_tot_26 = sum(r26.values())
        mgr_tot_25 = sum(r25.values())

        # Agrupar em grupos regionais conforme o gerente:
        #  - Bruno Tepedino: por estado (UF)
        #  - Leandro Marinho / Rafael Leiro: subregiões Norte/NE1/NE2/NE3/CO1/CO2
        #  - demais: distritos SAP resumidos (SP dividido em SP1/SP2/Interior)
        grp_26: dict[str, int] = defaultdict(int)
        grp_25: dict[str, int] = defaultdict(int)
        grp_label: dict[str, str] = {}
        for reg in all_regs:
            if mgr_u == "BRUNO TEPEDINO":
                key, label = _bruno_region_group(reg)
            elif mgr_u in ("LEANDRO MARINHO", "RAFAEL LEIRO"):
                key, label = _territorial_region_group(reg)
            else:
                key, label = _region_group(reg, sp_division)
            grp_26[key] += r26.get(reg, 0)
            grp_25[key] += r25.get(reg, 0)
            grp_label[key] = label

        # Leandro Marinho / Rafael Leiro: ordem fixa Norte/NE1/NE2/NE3/CO1/CO2
        # Lucas Gomes / Luciano Pereira: ordem fixa SP1/SP2/SP Interior (+ subtotal SP)/RJ/MG/ES
        # demais: volume 2026 desc
        if mgr_u in ("LEANDRO MARINHO", "RAFAEL LEIRO"):
            grp_keys_sorted = sorted(grp_label, key=lambda k: _TERRITORIAL_GROUP_ORDER.get(k, 999))
        elif mgr_u in ("LUCAS GOMES", "LUCIANO PEREIRA"):
            grp_keys_sorted = sorted(grp_label, key=lambda k: _SUDESTE_GROUP_ORDER.get(k, 999))
        else:
            grp_keys_sorted = sorted(grp_label, key=lambda k: -grp_26.get(k, 0))

        sp_keys_present = [k for k in grp_keys_sorted if k in _SUDESTE_SP_KEYS]
        last_sp_key = sp_keys_present[-1] if sp_keys_present else None

        reg_rows = ""
        for row_idx, key in enumerate(grp_keys_sorted):
            v26   = grp_26.get(key, 0)
            v25   = grp_25.get(key, 0)
            bg    = "background:#f7f7f5;" if row_idx % 2 else ""
            label = _SUDESTE_LABEL_OVERRIDE.get(key, grp_label[key]) if mgr_u in ("LUCAS GOMES", "LUCIANO PEREIRA") else grp_label[key]
            reg_rows += (
                f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
                f'<td style="padding:4px 8px 4px 14px;{BODY_FONT};font-size:11px;font-weight:bold">{label}</td>'
                f'<td style="padding:4px 8px;text-align:right;{TITLE_FONT};font-size:13px;font-weight:bold">{fmt(v26) if v26 else "&mdash;"}</td>'
                f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(v25) if v25 else "&mdash;"}</td>'
                f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px">{_pct_html(v26, v25)}</td>'
                f'</tr>'
            )
            if mgr_u in ("LUCAS GOMES", "LUCIANO PEREIRA") and key == last_sp_key:
                sp_tot_26 = sum(grp_26.get(k, 0) for k in _SUDESTE_SP_KEYS)
                sp_tot_25 = sum(grp_25.get(k, 0) for k in _SUDESTE_SP_KEYS)
                reg_rows += (
                    f'<tr style="background:#e6e6e2;border-top:1px solid #cccccc;border-bottom:2px solid #cccccc">'
                    f'<td style="padding:4px 8px 4px 14px;{BODY_FONT};font-size:9px;font-weight:bold;text-transform:uppercase;color:#555">Subtotal SP</td>'
                    f'<td style="padding:4px 8px;text-align:right;{TITLE_FONT};font-size:13px;font-weight:bold">{fmt(sp_tot_26) if sp_tot_26 else "&mdash;"}</td>'
                    f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(sp_tot_25) if sp_tot_25 else "&mdash;"}</td>'
                    f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px">{_pct_html(sp_tot_26, sp_tot_25)}</td>'
                    f'</tr>'
                )

        reg_blocks += (
            f'<td valign="top" style="padding-right:10px">'
            f'<div style="background:#111111;color:#ffffff;padding:4px 8px;{BODY_FONT};font-size:9px;font-weight:bold;letter-spacing:2px;text-transform:uppercase">{_mgr(mgr)}</div>'
            f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">'
            f'<thead><tr style="border-top:2px solid #111111;border-bottom:1px solid #dddddd">'
            f'<th style="padding:4px 8px;text-align:left;font-size:8px;font-weight:700;text-transform:uppercase;color:#333;{BODY_FONT}">Regi&atilde;o</th>'
            f'<th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;text-transform:uppercase;color:#333;{BODY_FONT}">{YTD_YEAR}</th>'
            f'<th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;text-transform:uppercase;color:#888;{BODY_FONT}">{LY_YEAR}</th>'
            f'<th style="padding:4px 8px;text-align:right;font-size:8px;font-weight:700;text-transform:uppercase;color:#333;{BODY_FONT}">YoY</th>'
            f'</tr></thead>'
            f'<tbody>{reg_rows}'
            f'<tr style="background:#eeeeee;border-top:2px solid #cccccc">'
            f'<td style="padding:4px 8px;{BODY_FONT};font-size:9px;font-weight:bold;text-transform:uppercase;color:#555">Subtotal</td>'
            f'<td style="padding:4px 8px;text-align:right;{TITLE_FONT};font-size:13px;font-weight:bold">{fmt(mgr_tot_26)}</td>'
            f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:11px;color:#888">{fmt(mgr_tot_25)}</td>'
            f'<td style="padding:4px 8px;text-align:right;{BODY_FONT};font-size:10px">{_pct_html(mgr_tot_26, mgr_tot_25)}</td>'
            f'</tr></tbody></table></td>'
        )

    sec05 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>{reg_blocks}</tr>
  </table>
</td></tr>"""

    # =========================================================================
    # §06 — POR MARCA (por segmento, lado a lado)
    # =========================================================================
    def _brand_table(seg):
        seg_brands = sorted(
            [b for b in set().union(*[set(v.keys()) for v in ytd26.values()])
             if BRAND_SEGMENT.get(b) == seg],
            key=lambda b: -sum(ytd26.get(m, {}).get(b, 0) for m in ytd26),
        )
        if not seg_brands:
            return ""
        s26  = sum(sum(ytd26.get(m, {}).get(b, 0) for m in ytd26) for b in seg_brands)
        s25y = sum(sum(ytd25y.get(m, {}).get(b, 0) for m in ytd25y) for b in seg_brands)
        hdr_bg = SEG_COLOR.get(seg, "#333333")
        rows = f"""
        <tr style="background:{hdr_bg}">
          <td colspan="4" style="padding:5px 8px;color:#ffffff;{BODY_FONT};font-size:10px;font-weight:bold;letter-spacing:2px;text-align:center">{seg} &mdash; {fmt(s26)} <span style="font-size:9px;font-weight:normal">({_pct_html_dark(s26, s25y, bold=False)})</span></td>
        </tr>
        <tr style="border-bottom:1px solid #cccccc;background:#f7f7f7">
          <th style="padding:3px 6px;text-align:left;font-size:7.5px;text-transform:uppercase;font-weight:700;color:#555;{BODY_FONT}">Marca</th>
          <th style="padding:3px 6px;text-align:right;font-size:7.5px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">{YTD_YEAR} YTD</th>
          <th style="padding:3px 6px;text-align:right;font-size:7.5px;text-transform:uppercase;font-weight:700;color:#888;{BODY_FONT}">{LY_YEAR} YTD</th>
          <th style="padding:3px 6px;text-align:right;font-size:7.5px;text-transform:uppercase;font-weight:700;color:#333;{BODY_FONT}">YoY</th>
        </tr>"""
        for j, brand in enumerate(seg_brands):
            v26  = sum(ytd26.get(m, {}).get(brand, 0) for m in ytd26)
            v25y = sum(ytd25y.get(m, {}).get(brand, 0) for m in ytd25y)
            if not v26 and not v25y:
                continue
            bg = "background:#f9f9f9;" if j % 2 else ""
            bname = BRAND_NAMES.get(brand, "")
            rows += (
                f'<tr style="{bg}border-bottom:1px solid #eeeeee">'
                f'<td style="padding:3px 6px 3px 12px;{BODY_FONT};font-size:10px;font-weight:bold;color:#444">'
                f'{brand}<span style="color:#aaa;font-weight:normal;font-size:8.5px;margin-left:4px">{bname}</span></td>'
                f'<td style="padding:3px 6px;text-align:right;{TITLE_FONT};font-size:11px">{fmt(v26)}</td>'
                f'<td style="padding:3px 6px;text-align:right;{BODY_FONT};font-size:10.5px;color:#888">{fmt(v25y) if v25y else "&mdash;"}</td>'
                f'<td style="padding:3px 6px;text-align:right;{BODY_FONT};font-size:10px">{_pct_html(v26, v25y)}</td>'
                f'</tr>'
            )
        rows += (
            f'<tr style="background:{hdr_bg};border-top:2px solid #111111">'
            f'<td style="padding:4px 6px;{BODY_FONT};font-size:10px;font-weight:bold;color:#fff">TOTAL</td>'
            f'<td style="padding:4px 6px;text-align:right;{TITLE_FONT};font-size:12px;font-weight:bold;color:#fff">{fmt(s26)}</td>'
            f'<td style="padding:4px 6px;text-align:right;{BODY_FONT};font-size:11px;color:#ccc">{fmt(s25y)}</td>'
            f'<td style="padding:4px 6px;text-align:right;{BODY_FONT};font-size:10px">{_pct_html_dark(s26, s25y)}</td>'
            f'</tr>'
        )
        return (
            f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;{BODY_FONT}">'
            f'<tbody>{rows}</tbody></table>'
        )

    brand_cols = ""
    for k, seg in enumerate(SEG_ORDER):
        pr = "padding-right:10px;" if k < len(SEG_ORDER) - 1 else ""
        brand_cols += f'<td valign="top" style="{pr}">{_brand_table(seg)}</td>'
        if k < len(SEG_ORDER) - 1:
            brand_cols += '<td width="1" style="background:#dddddd">&nbsp;</td>'

    sec06 = f"""
<tr><td style="padding:8px 34px 0 34px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>{brand_cols}</tr>
  </table>
</td></tr>"""

    # =========================================================================
    # Montar HTML completo
    # =========================================================================
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>ORDER SALES NEWSLETTER &mdash; {today_lbl}</title>
</head>
<body style="margin:0;padding:0;background:#e4e2dd">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#e4e2dd">
<tr><td align="center" style="padding:20px 12px">

<table role="presentation" width="800" cellpadding="0" cellspacing="0" border="0"
       style="width:800px;max-width:800px;background:#ffffff;border:1px solid #dddddd">

  <!-- MASTHEAD -->
  <tr><td style="padding:26px 34px 0 34px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td align="left" valign="bottom">
        <div style="{TITLE_FONT};font-weight:bold;font-size:28px;letter-spacing:8px;color:#111111;line-height:1">MARCOLIN</div>
        <div style="{TITLE_FONT};font-size:10px;letter-spacing:7px;color:#111111;padding-top:3px">EYEWEAR &#183; BRASIL</div>
      </td>
      <td align="right" valign="bottom">
        <div style="display:inline-block;background:#111111;color:#ffffff;{BODY_FONT};font-weight:bold;font-size:9px;letter-spacing:2px;padding:5px 11px">ORDER SALES NEWSLETTER</div>
        <div style="{BODY_FONT};font-size:10px;color:#777777;padding-top:7px">GERENTES &amp; CEO &#183; USO INTERNO</div>
      </td>
    </tr></table>
  </td></tr>
  <tr><td style="padding:12px 34px 0 34px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr><td style="border-bottom:3px solid #111111;font-size:0;line-height:0">&nbsp;</td></tr>
    </table>
  </td></tr>

  <!-- AVISO -->
  <tr><td style="padding:10px 34px 0 34px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #d8d8d8;background:#fafaf8"><tr>
      <td style="{TITLE_FONT};font-weight:bold;font-size:9px;letter-spacing:1px;color:#111111;white-space:nowrap;border-right:1px solid #cccccc;width:1px;padding:7px 12px">AVISO</td>
      <td style="{BODY_FONT};font-size:10px;color:#666666;line-height:1.4;padding:7px 12px">Dados referentes ao Ordenado SAP &mdash; Base D&minus;1 &middot; Valores sujeitos a atualiza&ccedil;&atilde;o.</td>
    </tr></table>
  </td></tr>

  <!-- METADADOS -->
  <tr><td style="padding:12px 34px 0 34px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td align="left" style="{BODY_FONT};font-size:10px;color:#555555">{today_lbl}</td>
      <td align="center" style="{BODY_FONT};font-size:10px;color:#555555">Semana {cur_week:02d} &middot; {month_lbl}</td>
      <td align="right" style="{BODY_FONT};font-size:10px;color:#555555">YTD: Jan&ndash;{today.strftime('%b %Y').upper()} &middot; {YTD_YEAR} vs {LY_YEAR}</td>
    </tr></table>
  </td></tr>

  <!-- TÍTULO -->
  <tr><td style="padding:22px 34px 0 34px">
    <div style="{TITLE_FONT};font-size:11px;letter-spacing:4px;text-transform:uppercase;color:#999999">ORDENADO &#183; JAN&ndash;{today.strftime('%b').upper()} {YTD_YEAR}</div>
    <div style="{TITLE_FONT};font-weight:bold;font-size:38px;letter-spacing:-1px;line-height:1;text-transform:uppercase;color:#111111;padding-top:3px">SALES ORDERS</div>
  </td></tr>

  {section_head("01", f"VISÃO GERAL — YTD &amp; HOJE")}
  {sec01}

  {section_head("02", f"PERFORMANCE POR GERENTE — YTD {YTD_YEAR}")}
  {sec02}

  {section_head("03", f"MÊS × MÊS — {month_lbl} (semanas)")}
  {sec03}

  {section_head("04", f"SEMANA × SEMANA — Sem. {cur_week:02d} &middot; {today.strftime('%d/%m/%Y')}")}
  {sec04}

  {section_head("05", "REGIONAL — YTD por Gerente")}
  {sec05}

  {section_head("06", "POR SEGMENTO &amp; MARCA — YTD")}
  {sec06}

  <!-- RODAPÉ -->
  <tr><td style="padding:28px 34px 26px 34px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr><td style="border-top:3px solid #111111;font-size:0;line-height:0;padding-bottom:10px">&nbsp;</td></tr>
    </table>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td align="left" style="{BODY_FONT};font-size:9px;color:#999999">MARCOLIN EYEWEAR BRASIL &#183; PROJETO ODIN</td>
      <td align="center" style="{BODY_FONT};font-size:9px;color:#999999">Fonte: SAP Ordenado &mdash; {today_str}</td>
      <td align="right" style="{BODY_FONT};font-size:9px;color:#999999">Confidencial &#183; N&atilde;o Distribuir</td>
    </tr></table>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""

    return html


# ── Intro HTML ─────────────────────────────────────────────────────────────────

def build_intro(gen_date):
    today_str = gen_date.strftime("%d/%m/%Y")
    month_lbl = gen_date.strftime("%b %Y").title()
    return f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>
<body style="{BODY_FONT};font-size:13px;color:#222;line-height:1.7;margin:0;padding:0;background:#fff">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff"><tr>
<td style="padding:24px 32px 0 32px;max-width:800px">

<p style="margin:0 0 14px 0">Prezados,</p>
<p style="margin:0 0 14px 0">
  Segue a <strong>Order Sales Newsletter</strong> com o desempenho consolidado de pedidos
  referente ao período <strong>Jan&ndash;{month_lbl} {YTD_YEAR} vs {LY_YEAR}</strong>,
  posi&ccedil;&atilde;o de <strong>{today_str}</strong> (base D&minus;1).
</p>
<p style="margin:0 0 6px 0">O relat&oacute;rio contempla:</p>
<ul style="margin:0 0 14px 22px;padding:0;color:#444;font-size:12.5px">
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;01</strong> &mdash; Vis&atilde;o Geral: YTD + indicadores do dia</li>
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;02</strong> &mdash; Performance por Gerente (YTD 2026 vs 2025)</li>
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;03</strong> &mdash; M&ecirc;s &times; M&ecirc;s (semanas de {month_lbl})</li>
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;04</strong> &mdash; Semana &times; Semana (dias da semana {gen_date.isocalendar()[1]:02d})</li>
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;05</strong> &mdash; Regional por Gerente (YTD)</li>
  <li style="margin-bottom:3px"><strong>&sect;&nbsp;06</strong> &mdash; Por Segmento &amp; Marca (YTD)</li>
</ul>
<p style="margin:0 0 20px 0;color:#555;font-size:12px;font-style:italic">
  Para d&uacute;vidas, contate a &aacute;rea de <strong style="color:#111">Sales Intelligence</strong>.
</p>
<p style="margin:0 0 3px 0">Atenciosamente,</p>
<p style="margin:0 0 28px 0;font-weight:700;color:#111">Marcolin Sales Intelligence</p>
<hr style="border:none;border-top:2px solid #e0e0e0;margin:0 0 24px 0">
</td></tr></table>"""


# ── Outlook sender ─────────────────────────────────────────────────────────────

BCC_RECIPIENTS = [
    "lmarinho@marchon.com.br",
    "rafael.dominguez@marchon.com.br",
    "lpereira@marcolin.com",
    "lgomes@marchon.com.br",
    "btepedino@marcolin.com",
    "mpinheiro@marcolin.com",
    "cmalvao@marcolin.com",
    "mjunior@marcolin.com",
    "pgoncalves@marcolin.com"
]

# SEND_FROM_ACCOUNT = "comercialbr@marcolin.com"  # TODO: reativar quando o TI adicionar essa conta ao perfil do Outlook


def send_outlook(subject, html_body, recipients):
    try:
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        mail            = outlook.CreateItem(0)
        # TODO: reativar quando o envio pela caixa compartilhada for liberado
        mail.SentOnBehalfOfName = "reporting_br@marcolin.com"
        mail.To         = "; ".join(recipients)
        mail.BCC        = "; ".join(BCC_RECIPIENTS)
        mail.Subject    = subject
        mail.HTMLBody   = html_body

        # for account in outlook.Session.Accounts:
        #     if account.SmtpAddress.lower() == SEND_FROM_ACCOUNT.lower():
        #         mail.SendUsingAccount = account
        #         break
        # else:
        #     print(f"  AVISO — conta {SEND_FROM_ACCOUNT} não encontrada no perfil do Outlook; usando conta padrão.", flush=True)

        mail.Send()  # Abre o e-mail no Outlook para revisão e envio manual
        print(f"  OK — e-mail preparado para: {', '.join(recipients)}", flush=True)
        return True
    except Exception as e:
        print(f"  ERRO ao preparar e-mail: {e}", flush=True)
        return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    run_ts    = datetime.today()
    gen_date  = datetime.combine(_TODAY, datetime.min.time())  # D-1 — dados fechados do dia anterior
    today_str = gen_date.strftime("%d/%m/%Y")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("=" * 60)
    print("  MARCOLIN — ORDER SALES NEWSLETTER")
    print(f"  Execução: {run_ts.strftime('%d/%m/%Y %H:%M')}  |  Referência (D-1): {today_str}")
    print("=" * 60)

    print("\nCarregando mapa chave→gerente...", flush=True)
    key_to_manager = load_key_to_manager()
    print(f"  {len(key_to_manager):,} chaves carregadas.", flush=True)

    print("Carregando pedidos...", flush=True)
    data = load_orders(key_to_manager)
    tot  = sum(sum(v.values()) for v in data["ytd26"].values())
    print(f"  YTD 2026: {tot:,} peças.", flush=True)

    print("Gerando HTML...", flush=True)
    html = build_html(gen_date, data)

    fname = f"orders_{gen_date.strftime('%Y%m%d')}.html"
    path  = os.path.join(OUTPUT_DIR, fname)
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Salvo: {path}", flush=True)

    intro      = build_intro(gen_date)
    email_html = intro + html
    subject    = f"[Marcolin] Order Sales Newsletter — {today_str}"
    send_outlook(subject, email_html, RECIPIENTS)

    print()
    print("=" * 60)
    print(f"  CONCLUÍDO — {path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
