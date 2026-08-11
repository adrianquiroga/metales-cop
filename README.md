# 🪙 Módulo de Seguimiento Diario de Metales Preciosos (Oro y Plata en COP)

Este módulo automatiza la consulta, cálculo y almacenamiento de la estadística diaria del precio en gramos del **Oro** y la **Plata** en Pesos Colombianos (COP), cubriendo el historial completo desde el **1 de Enero de 2025** hasta la fecha actual.

---

## 🌐 Enlace Web Público (GitHub Pages)

👉 **Sitio Web Interactivo:** [https://adrianquiroga.github.io/metales-cop/](https://adrianquiroga.github.io/metales-cop/)

---

## 📁 Archivos del Módulo

```
c:\PROYECTOS\PRESUPUESTO\metales\
├── .github/workflows/
│   └── actualizar.yml          # 🤖 Automatización en la nube (GitHub Actions a las 6:00 PM)
├── index.html                  # 🌐 Interfaz Web Interactiva (Gráfica al 100% de ancho con Tooltips)
├── data.js                     # 📊 Base de datos JavaScript exportada automáticamente
├── data.json                   # 📄 Base de datos JSON para APIs o consumo externo
├── actualizar_metales.py       # 🐍 Script principal de actualización en Python
├── actualizar_metales.bat      # ⚡ Ejecutable automático en 1 Clic (Python + Publicar en GitHub)
├── Precios_Metales_COP.xlsx    # 📈 Libro de Excel con Dashboard y datos diarios
└── README.md                   # 📝 Documentación del módulo
```

---

## 🤖 Automatización Autónoma en la Nube (GitHub Actions)

El archivo [`.github/workflows/actualizar.yml`](file:///c:/PROYECTOS/PRESUPUESTO/metales/.github/workflows/actualizar.yml) configura un robot en los servidores de GitHub que se ejecuta automáticamente de **lunes a viernes a las 6:00 PM (hora Colombia)**.

### Lo que hace el robot en la nube:
1. Enciende un servidor virtual Ubuntu en GitHub.
2. Instala Python y `openpyxl`.
3. Ejecuta `python actualizar_metales.py` consultando Yahoo Finance.
4. Actualiza `Precios_Metales_COP.xlsx`, `data.js` y `data.json`.
5. Realiza `git commit` y `git push` automáticamente.
6. GitHub Pages renueva la web pública [https://adrianquiroga.github.io/metales-cop/](https://adrianquiroga.github.io/metales-cop/) sin que tengas que tocar nada.

---

## ⚡ Opción Manual en 1 Clic (Local)

También puedes ejecutar manualmente en cualquier momento haciendo **doble clic en [`actualizar_metales.bat`](file:///c:/PROYECTOS/PRESUPUESTO/metales/actualizar_metales.bat)**.

---

## 📐 Fórmula de Conversión a Pesos Colombianos por Gramo

$$\text{Precio Oro (COP/g)} = \left( \frac{\text{Precio Oro USD/oz}}{31.1034768} \right) \times \text{TRM (USD/COP)}$$

$$\text{Precio Plata (COP/g)} = \left( \frac{\text{Precio Plata USD/oz}}{31.1034768} \right) \times \text{TRM (USD/COP)}$$
