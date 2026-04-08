# Timeline PPT Generator

Genera presentaciones PowerPoint con líneas de tiempo a partir de una lista de eventos en JSON. Los eventos se ordenan cronológicamente y se distribuyen en slides con un diseño visual profesional.

## Requisitos

- Python 3.10+
- Dependencias: `python-pptx`

```bash
pip install -r requirements.txt
```

## Uso

### Comando básico

```bash
python timeline_generator.py --events eventos.json --output mi_timeline.pptx
```

### Con un PPT existente como base (agrega slides al final)

```bash
python timeline_generator.py --template mi_plantilla.pptx --events eventos.json --output resultado.pptx
```

### Opciones

| Argumento          | Descripción                                                      | Default                |
|--------------------|------------------------------------------------------------------|------------------------|
| `--events`         | **(Requerido)** Ruta al archivo JSON con los eventos.            | —                      |
| `--output`         | Ruta del archivo PPT de salida.                                  | `timeline_output.pptx` |
| `--template`       | PPT existente donde se agregarán los slides de timeline al final.| `None`                 |
| `--title`          | Título que aparece en la barra superior de cada slide.           | `Línea de Tiempo`      |
| `--max-per-slide`  | Cantidad máxima de eventos por slide.                            | `8`                    |

## Formato del JSON de eventos

El archivo JSON debe ser una lista de objetos. Cada evento necesita al menos `date` y `title`:

```json
[
  {
    "date": "2023-01-15",
    "display_date": "15 Ene 2023",
    "title": "Lanzamiento del producto",
    "description": "Primera versión disponible al público."
  },
  {
    "date": "2023-06",
    "title": "Expansión a nuevos mercados"
  }
]
```

### Campos

| Campo          | Requerido | Descripción                                                                 |
|----------------|-----------|-----------------------------------------------------------------------------|
| `date`         | Sí        | Fecha del evento. Formatos: `YYYY-MM-DD`, `YYYY-MM`, o `YYYY`.             |
| `title`        | Sí        | Título del evento (texto principal).                                        |
| `display_date` | No        | Texto para mostrar como fecha en el slide (si no se pone, se usa `date`).  |
| `description`  | No        | Descripción adicional debajo del título.                                    |

## Ejemplo

Hay un archivo de ejemplo incluido:

```bash
python timeline_generator.py --events example_events.json --output demo.pptx --title "Historia del Proyecto"
```

Esto genera un PPT con 10 eventos distribuidos en 2 slides.

## Diseño generado

Cada slide incluye:
- Barra de título azul en la parte superior
- Línea horizontal con puntos para cada evento
- Fechas y descripciones alternando arriba/abajo para mejor legibilidad
- Flecha al final indicando continuidad temporal
- Si hay más eventos que el máximo por slide, se generan slides adicionales numerados
