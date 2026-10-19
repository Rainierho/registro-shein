# -*- coding: utf-8 -*-
"""Piezas compartidas por app.py, catalogo.py y enlaces.py, para que coincidan siempre."""
import re
import unicodedata

FOTO_PROPIA = "foto_propia"     # formato viejo: una foto subida, ligada al nombre del producto


def slug(nombre):
    """ANILLO DE DRAGÓN -> anillo-de-dragon (la dirección de cada producto)."""
    s = unicodedata.normalize("NFD", str(nombre or "").lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def fotos_de(valor):
    """Celda FOTO -> fotos en orden: enlaces, «propia:<id>» o el formato viejo."""
    salida = []
    for parte in str(valor or "").split("|"):
        q = parte.strip()
        if q.startswith("http") or q.startswith("propia:") or q == FOTO_PROPIA:
            salida.append(q)
    return salida
