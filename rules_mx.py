"""
rules_mx.py — Reglas de mapeo processor por entidad México.

Estructura de cada regla:
    (keyword_en_descripcion, nombre_processor, account_id, account_code)

Orden importa: la primera regla que coincida gana.
"""

from typing import NamedTuple


class Rule(NamedTuple):
    keyword: str
    processor: str
    account_id: str
    account_code: str


# ── DLOCAL MEXICO (MX03) ──────────────────────────────────────────────────────

RULES_MX_DLOCAL: list[Rule] = [
    Rule("D LOCAL",                   "Banorte",         "1151995728",  "AA370"),
    Rule("EVOPAYMX",                  "EVO MPGs",         "70177702216", "AA376"),
    Rule("NET CARD SALES",            "EVO MPGs",         "70177702216", "AA376"),
    Rule("ABONO VENTAS",              "Hey Banregio",     "220881970023","AA374"),
    Rule("MERCADO PAGOREFERENCIA",    "Mercadopago",      "1151995728",  "AA370"),
    Rule("MP AGREGADOR",              "Mercadopago",      "1151995728",  "AA370"),
    Rule("OPENMX",                    "Openpay",          "116158803",   "AA375"),
    Rule("CADENA COMERCIAL OXXO SA",  "OXXO Pay",         "1151995728",  "AA370"),
    Rule("LIQ DLOCALMEXICO",          "Kushki",           "1151995728",  "AA370"),
    Rule("LIQ DLOCAL",                "Kushki",           "1151995728",  "AA370"),
]

# Account-code fallback para Dlocal Mexico cuando las reglas de keyword no matchean
ACCOUNT_MAP_DLOCAL: dict[str, str] = {
    "AA370": "Banorte",
    "AA374": "Hey Banregio",
    "AA375": "Openpay",
    "AA376": "EVO MPGs",
    "AA566": "Banco Santander",   # USD account — treasury, filtrar
    "AA639": "Banco Santander",
    "AA640": "EVO MPGs",
    "AB072": "Openpay",
    "AB184": "STP",
    "AB279": "BBVA Bancomer",
}

# ── DEMERGE MEXICO (MX02) ──────────────────────────────────────────────────────

RULES_MX_DEMERGE: list[Rule] = [
    Rule("D LOCAL",                   "Banorte",         "1011320992",  "AA350"),
    Rule("EVOPAYMX",                  "EVO MPGs",         "70137911173", "AA358"),
    Rule("NET CARD SALES",            "EVO MPGs",         "70137911173", "AA358"),
    Rule("ABONO VENTAS",              "Hey Banregio",     "220881930013","AA354"),
    Rule("MP AGREGADOR",              "Mercadopago",      "1011320992",  "AA350"),
    Rule("MERCADO PAGOREFERENCIA",    "Mercadopago",      "1011320992",  "AA350"),
    Rule("OPENMX",                    "Openpay",          "111698419",   "AA356"),
    Rule("OPENMX",                    "Openpay_paynet",   "113735494",   "AA357"),
    Rule("CADENA COMERCIAL OXXO SA",  "OXXO Pay",         "1011320992",  "AA350"),
    Rule("LIQ DEMEREGE BIG PLAYERS",  "Kushki",           "1011320992",  "AA350"),
    Rule("LIQ DLOCAL",                "Kushki",           "1011320992",  "AA350"),
]

ACCOUNT_MAP_DEMERGE: dict[str, str] = {
    "AA350": "Banorte",
    "AA354": "Hey Banregio",
    "AA356": "Openpay",
    "AA357": "Openpay_paynet",
    "AA358": "EVO MPGs",
}

# ── REGISTRO DE ENTIDADES ──────────────────────────────────────────────────────

ENTITY_CONFIG = {
    "Dlocal Mexico": {
        "rules":       RULES_MX_DLOCAL,
        "account_map": ACCOUNT_MAP_DLOCAL,
        "company_ids": ["MX03", "dlocal mexico"],  # substrings del campo Company en Kyriba
        "valid_agents": ["dlocal mexico", "dlocal technologies"],
    },
    "Demerge Mexico": {
        "rules":       RULES_MX_DEMERGE,
        "account_map": ACCOUNT_MAP_DEMERGE,
        "company_ids": ["MX02", "demerge mexico"],
        "valid_agents": ["demerge mexico"],
    },
}

# ── PALABRAS CLAVE TREASURY / INTERNAL ────────────────────────────────────────
# Movimientos que NO son payins y deben excluirse del análisis.

TREASURY_KEYWORDS = [
    "TRASPASO",
    "SPEI ENVIADO",
    "SPEI A BANORTE",
    "SPEI A SANTANDER",
    "SWEEP",
    "BARRIDO",
    "INTERCOMPANY",
    "INTERNAL TRANSFER",
    "ENV TRANSF",
    "C-V MONEDA EXTRANJERA",     # FX trade
    "CARGO POR TRANSFERENCIA",
    "COMISION DE MENSAJE",
    "COMIS TAR",
    "COBRO IMP INST",
    "COBRO COMIS",
    "DEBIT BY INSTRUCTION",      # Citi fee rows
    "CARGO ACL AFIL",            # card fee rows in Banregio
]

# Cuentas cuya naturaleza es 100% treasury/FX — se excluyen siempre como payins
TREASURY_ACCOUNTS = {
    "AA566",   # Citi USD — Dlocal Mexico
    "AB072",   # Banorte USD — Dlocal Mexico
    "AB279",   # BBVA Bancomer — treasury outflows
}

# ── NORMALIZACIÓN DE NOMBRES DE PROCESSOR ─────────────────────────────────────

PROCESSOR_ALIASES: dict[str, str] = {
    # Banorte
    "banorte": "Banorte",
    "banco banorte": "Banorte",
    "banorte mx": "Banorte",
    "dlocal banorte": "Banorte",
    "demerge banorte": "Banorte",
    # BBVA
    "bbva": "BBVA Bancomer",
    "bbva bancomer": "BBVA Bancomer",
    "banco bancomer": "BBVA Bancomer",
    "bancomer": "BBVA Bancomer",
    # Santander
    "santander": "Banco Santander",
    "banco santander": "Banco Santander",
    # Citi
    "citi": "CITI",
    "citibanamex": "CITI",
    "banamex": "CITI",
    # STP
    "stp": "STP",
    # EVO
    "evo mpgs": "EVO MPGs",
    "evo mpgs mx": "EVO MPGs",
    "evopaymx": "EVO MPGs",
    "evo payments": "EVO MPGs",
    "evo": "EVO MPGs",
    # Banregio
    "hey banregio": "Hey Banregio",
    "banregio": "Hey Banregio",
    "banco banregio": "Hey Banregio",
    # Mercadopago
    "mercadopago": "Mercadopago",
    "mercado pago": "Mercadopago",
    "mercado pago mx": "Mercadopago",
    "mercado pago referencia": "Mercadopago",
    "mp": "Mercadopago",
    "mp agregador": "Mercadopago",
    # Openpay
    "openpay": "Openpay",
    "openpay mx": "Openpay",
    "openpay spei": "Openpay",
    "openpay_spei": "Openpay",
    "open pay": "Openpay",
    "dlocal openpay": "Openpay",
    # Openpay Paynet
    "openpay paynet": "Openpay_paynet",
    "openpay_paynet": "Openpay_paynet",
    "paynet": "Openpay_paynet",
    "paynet mx": "Openpay_paynet",
    # OXXO
    "oxxo pay": "OXXO Pay",
    "oxxopay": "OXXO Pay",
    "oxxo": "OXXO Pay",
    "oxxo mx": "OXXO Pay",
    # Arcus
    "arcus": "Arcus",
    "arcus mx": "Arcus",
    "dlocal arcus": "Arcus",
    # Kushki
    "kushki": "Kushki",
    "kushki mexico": "Kushki",
    "kushki mx": "Kushki",
    # Dlocal Technologies (sub-entidad dentro de Dlocal Mexico)
    "dlocal technologies": "Dlocal Technologies",
}
