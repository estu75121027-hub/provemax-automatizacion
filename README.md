# ProveMax - Automatización con Python

Mini-sistema local que convierte la página de la tienda **ProveMax** en una app con backend en **Python + Flask**: el catálogo se sirve por API y cada pedido se registra automáticamente en un archivo Excel.

## Características

- Catálogo de productos en `productos.json` (12 productos).
- API REST con Flask:
  - `GET /api/productos` → devuelve el catálogo en JSON.
  - `POST /api/pedidos` → valida, calcula el total y guarda el pedido en `pedidos.xlsx`.
  - `GET /api/pedidos` → resumen de pedidos registrados.
- Carrito de compras en la web con persistencia en `localStorage`.
- Envío del pedido por WhatsApp con número correlativo (`P-0001`, `P-0002`, ...).
- Presentación del proyecto (Semana Tecnológica) en `presentacion.html`.

## Estructura

```
Analisis de datos/
├── app.py            ← Servidor Flask (API + lógica de pedidos)
├── productos.json    ← Catálogo de productos
├── pedidos.xlsx      ← Se genera al primer pedido (historial)
├── templates/
│   └── index.html    ← Tienda conectada a la API
└── presentacion.html ← Presentación del proyecto
```

## Requerimientos

- Python 3.8+ (probado en 3.12)
- Flask 3.1.3
- openpyxl 3.1.5

## Instalación y uso

```bash
pip install flask openpyxl
python app.py
# abrir http://localhost:5000
```

## Autor

Edson — Proyecto individual, Semana Tecnológica.
