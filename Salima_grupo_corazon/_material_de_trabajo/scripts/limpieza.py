# -*- coding: utf-8 -*-
"""Limpia de las fichas las notas técnicas sobre cómo se obtuvo cada texto."""
import re
REGLAS = [
    (r'la pasarela Wiley Scholar Gateway', 'la plataforma de Wiley'),
    (r'del conector Wiley Scholar Gateway', 'de la plataforma de Wiley'),
    (r'la herramienta Wiley Scholar Gateway', 'la plataforma de Wiley'),
    (r'Wiley Scholar Gateway', 'la plataforma de Wiley'),
    (r'con la herramienta WebFetch', 'directamente'),
    (r', el intento por la API de Elsevier no prosperó', ''),
    (r',?\s*(?:desde|en) (?:este|el) entorno', ''),
    (r'la verificación anti-?(?:bots|robots)', 'el control de acceso'),
    (r'una verificación anti-?(?:bots|robots)', 'un control de acceso'),
    (r'verificación anti-?(?:bots|robots)', 'control de acceso'),
    (r'anti-?(?:bots|robots)', 'de control de acceso'),
    (r'están bloquead[oa]s', 'tienen el acceso restringido'),
    (r'está bloquead[oa]', 'tiene el acceso restringido'),
    (r'bloquea las descargas automáticas', 'no permite la descarga directa'),
    (r'bloquea el acceso automático', 'restringe el acceso'),
    (r'bloquead[oa]s?', 'de acceso restringido'),
    (r'descargas automáticas', 'descargas directas'),
    (r'descarga automática', 'descarga directa'),
    (r'acceso automático', 'acceso directo'),
    (r'El borrador del apartado', 'La primera versión del apartado'),
    (r'El borrador', 'La primera versión del apartado'),
    (r'del borrador', 'de la primera versión del apartado'),
    (r'el borrador', 'la primera versión del apartado'),
    (r'  +', ' '),
]
def limpiar(s):
    if not isinstance(s, str): return s
    for a, b in REGLAS: s = re.sub(a, b, s)
    return s.strip()
