---

# Kardex Viewer

![Version](https://img.shields.io/badge/version-1.0.0-orange)
![Python](https://img.shields.io/badge/python-3.8+-blue)
![Status](https://img.shields.io/badge/status-en%20desarrollo-yellow)

Aplicación web desarrollada con **Python**, **Pandas** y **Streamlit** para la visualización y exportación de registros de inventario en formato Kardex.

> ⚠️ Primera versión. Proyecto en desarrollo activo.

## Requisitos

```bash
pip install streamlit pandas openpyxl
```

## Uso

```bash
streamlit run app_kardex_viewer.py
```

## Funcionalidades

- Carga múltiples archivos Excel (.xlsx) con estructura Kardex
- Visualiza todos los movimientos en una sola tabla unificada
- Filtra por año y mes
- Exporta los datos a Excel respetando la estructura y formatos originales
- Nombre personalizable para el archivo descargado

---
