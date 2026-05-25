"""
mapping.py — Lógica de asignación de Processors para México.

Responsabilidades:
  - Aplicar reglas keyword + account_id / account_code
  - Normalizar nombres de processor
  - Fallback por account code conocido
  - Parsing de mapeos manuales del usuario
"""

from __future__ import annotations

import re
from typing import Optional

import pandas as pd

from rules_mx import ENTITY_CONFIG, PROCESSOR_ALIASES, Rule


# ── NORMALIZACIÓN ──────────────────────────────────────────────────────────────

def normalize_processor(name: str) -> Optional[str]:
    """Devuelve el nombre canónico de un processor dado un alias."""
    if not isinstance(name, str) or not name.strip():
        return None
    clean = str(name).strip().lower().replace("_", " ").replace("-", " ")
    return PROCESSOR_ALIASES.get(clean, str(name).strip())


# ── PARSING DE MAPEO MANUAL ────────────────────────────────────────────────────

def parse_manual_mapping(text: str) -> dict[str, str]:
    """
    Convierte texto libre 'AA370=Banorte\\nAA376=EVO MPGs' en un dict.
    Claves en mayúsculas (account codes).
    """
    mapping: dict[str, str] = {}
    for line in str(text).splitlines():
        line = line.strip()
        if not line or "=" not in line:
            continue
        account, processor = line.split("=", 1)
        mapping[account.strip().upper()] = processor.strip()
    return mapping


# ── ENTITY MAPPER ──────────────────────────────────────────────────────────────

class EntityMapper:
    """
    Encapsula toda la lógica de mapeo de processors para una entidad.

    Uso:
        mapper = EntityMapper("Dlocal Mexico")
        processor = mapper.assign_kyriba_row(row)
        processor = mapper.normalize(raw_name)
    """

    def __init__(self, entity_name: str, manual_map: dict[str, str] | None = None):
        if entity_name not in ENTITY_CONFIG:
            raise ValueError(f"Entidad desconocida: {entity_name!r}")

        cfg = ENTITY_CONFIG[entity_name]
        self.entity_name: str          = entity_name
        self.rules: list[Rule]         = cfg["rules"]
        self.account_map: dict[str,str]= cfg["account_map"]
        self.valid_agents: list[str]   = cfg["valid_agents"]
        self.manual_map: dict[str,str] = manual_map or {}

        # Compilar pattern de keywords para match rápido
        self._rule_patterns = [
            (re.compile(re.escape(r.keyword), re.IGNORECASE), r)
            for r in self.rules
        ]

    # ── Kyriba ────────────────────────────────────────────────────────────────

    def assign_kyriba_row(self, row: "pd.Series") -> Optional[str]:
        """
        Asigna un processor a una fila Kyriba.

        Orden de prioridad:
            1. Mapeo manual del usuario (Account code → Processor)
            2. Reglas keyword + account_id / account_code
            3. Fallback account_map
        """
        code = str(row.get("Account code", "")).strip().upper()
        account_id = str(row.get("Account ID", "")).strip()
        text = str(row.get("Text to match", "")).upper()

        # 1. Manual override
        if code in self.manual_map:
            return normalize_processor(self.manual_map[code])

        # 2. Reglas keyword + account
        for pattern, rule in self._rule_patterns:
            if pattern.search(text):
                if (
                    account_id == str(rule.account_id).strip()
                    or code == str(rule.account_code).strip().upper()
                ):
                    return rule.processor

        # 3. Fallback por account code
        fallback = self.account_map.get(code)
        if fallback:
            return fallback

        return None

    # ── Payins ────────────────────────────────────────────────────────────────

    def normalize(self, raw_name: str) -> Optional[str]:
        """Normaliza un nombre de processor crudo (de archivo Payins)."""
        return normalize_processor(raw_name)

    # ── Validación de collection agent ────────────────────────────────────────

    def is_valid_agent(self, agent_str: str) -> bool:
        text = str(agent_str).strip().lower()
        return any(a in text for a in self.valid_agents)
