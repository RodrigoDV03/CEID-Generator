# CEID-Generator

## Nueva arquitectura

```text
CEID-Generator/
├── app.py
├── core/
│   ├── planillas/
│   ├── fases/
│   ├── correos/
│   └── control_pagos/
├── ui/
│   ├── layout.py
│   ├── components.py
│   └── views/
├── services/
├── utils/
├── assets/
│   ├── iconos/
│   └── imagenes/
├── data/
│   ├── firmas/
│   └── modelos/
└── requirements.txt
```

## Responsabilidades

- core: logica de negocio y procesamiento sin dependencias de UI.
- ui: vistas y componentes visuales de CustomTkinter.
- services: capa de orquestacion entre UI y core.
- utils: constantes y helpers transversales.
- data: firmas y plantillas de documentos.
- assets: recursos visuales del proyecto.

## Punto de entrada

Ejecutar:

```bash
python app.py
```

`Generador.py` se mantiene como alias temporal para compatibilidad y redirige a `app.py`.
