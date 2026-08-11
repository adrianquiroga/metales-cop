# 🪙 Módulo de Seguimiento Diario de Metales Preciosos (Oro y Plata en COP)

Este módulo automatiza la consulta, cálculo y almacenamiento de la estadística diaria del precio en gramos del **Oro** y la **Plata** en Pesos Colombianos (COP), cubriendo el historial completo desde el **1 de Enero de 2025** hasta la fecha actual.

---

## 📁 Archivos del Módulo

```
c:\PROYECTOS\PRESUPUESTO\metales\
├── index.html                  # 🌐 Interfaz Web Interactiva (Gráfica al 100% de ancho con Tooltips)
├── data.js                     # 📊 Base de datos JavaScript exportada automáticamente
├── data.json                   # 📄 Base de datos JSON para APIs o consumo externo
├── actualizar_metales.py       # 🐍 Script principal de actualización en Python
├── actualizar_metales.bat      # ⚡ Ejecutable rápido para Windows
├── Precios_Metales_COP.xlsx    # 📈 Libro de Excel con Dashboard y datos diarios
└── README.md                   # 📝 Documentación del módulo
```

---

## 🌐 1. Interfaz Web Interactiva (`index.html`)

Diseñada para resolver las limitaciones estáticas de Excel (aprovechamiento del 100% del eje X e interacción al pasar o hacer clic sobre los puntos).

### Características Principales:
- **Aprovechamiento total del eje X:** Al filtrar por *Últimos 30 Días*, *90 Días*, *2025*, *2026* o *Todo*, la curva se expande al **100% del ancho del gráfico**, eliminando espacios vacíos a los lados.
- **Tooltips Interactivos:** Al pasar o hacer clic en cualquier punto del gráfico, muestra una tarjeta emergente con:
  - Fecha exacta en español (*Ej. Jueves, 9 de Abril de 2026*).
  - Precio exacto formateado en COP (*Ej. $ 442.950,25 / g*).
  - Variación porcentual diaria.
- **Selector de Activo:** Botones para alternar entre Oro, Plata, TRM o Comparativo en Doble Eje.
- **Visualización Directa:** Se puede abrir haciendo **doble clic en `index.html`** directamente desde el explorador de archivos.

---

## 📈 2. Libro de Excel (`Precios_Metales_COP.xlsx`)

- **Solución al problema de Leyenda:** Se corrigió el parámetro `from_rows=False` en las series de `openpyxl`, eliminando las 400 entradas duplicadas en la leyenda de las gráficas.
- **Marcadores de datos:** Se activaron marcadores en los puntos de la línea para permitir clic sobre los valores.
- **Filtro Nativo de Excel:** En la pestaña `Precios Diarios` se incluye **AutoFiltro activo**. Al filtrar la columna `Fecha` directamente en Excel, la gráfica de Excel ajusta el eje X a las filas visibles.

---

## 🚀 Modo de Uso

1. **Para ver la Gráfica Interactiva Web:**
   - Haz doble clic sobre [`index.html`](file:///c:/PROYECTOS/PRESUPUESTO/metales/index.html).

2. **Para Actualizar Precios y Gráficas (Diario):**
   - Haz doble clic sobre [`actualizar_metales.bat`](file:///c:/PROYECTOS/PRESUPUESTO/metales/actualizar_metales.bat) o ejecuta:
     ```bash
     python metales/actualizar_metales.py
     ```
   *(Esto actualizará automáticamente tanto el archivo Excel como la aplicación web `index.html`)*.
