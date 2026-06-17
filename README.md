# Generador de Proformas — Nutricereales de Panamá

App en Streamlit que llena la plantilla `Nitri01.xlsx` (proforma) preservando formato, fórmulas y bordes. Catálogo de 107 productos en la `Hoja2`.

## Funcionalidad

- Búsqueda por código o nombre, filtro por categoría.
- Datos de cliente, entrega y contacto.
- Hasta 46 líneas de producto con UM auto-detectado del catálogo.
- Subtotal, descuento y total en USD calculados en vivo.
- Genera el archivo `.xlsx` listo para descargar, sin tocar la plantilla.
- Autoincremento del número de proforma (K3) opcional.

## Ejecutar local

```bash
pip install -r requirements.txt
streamlit run proforma_app.py
```

## Deploy en Streamlit Cloud

1. Sube el repo a GitHub.
2. Entra a [share.streamlit.io](https://share.streamlit.io) y conecta el repo.
3. Main file: `proforma_app.py`. Python: 3.11.

## Archivos

- `proforma_app.py` — app principal.
- `Nitri01.xlsx` — plantilla + catálogo.
- `requirements.txt` — dependencias.
