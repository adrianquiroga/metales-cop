# 🪙 Módulo de Seguimiento Diario de Metales Preciosos (Oro y Plata en COP)

Este módulo automatiza la consulta, cálculo y almacenamiento de la estadística diaria del precio en gramos del **Oro** y la **Plata** en Pesos Colombianos (COP), cubriendo el historial completo desde el **1 de Enero de 2025** hasta la fecha actual.

---

## 🌐 Enlace Web Público (GitHub Pages)

👉 **Sitio Web Interactivo:** [https://adrianquiroga.github.io/metales-cop/](https://adrianquiroga.github.io/metales-cop/)

---

## ⚡ Botón Administrador: Actualización Spot en Vivo (1 Clic)

Se implementó el botón destacado:

👉 **`⚡ Solicitar Precio Spot en Vivo (1 Clic)`**

### ¿Cómo funciona la privacidad del botón?
1. **Para Visitantes Públicos:** Al abrir [https://adrianquiroga.github.io/metales-cop/](https://adrianquiroga.github.io/metales-cop/), la interfaz se muestra limpia sin el botón de administrador.
2. **Para Ti (Administrador):** 
   - Puedes abrir la web con el enlace privado: **`https://adrianquiroga.github.io/metales-cop/?admin=1`**
   - O hacer clic en **`⚙️ Modo Administrador`** en la parte inferior de la página.
   - El botón amarillo **`⚡ Solicitar Precio Spot en Vivo (1 Clic)`** aparecerá en el encabezado.
   - Al tocarlo en tu celular, te abrirá GitHub con el botón `Run workflow` directo para disparar la actualización en vivo.

---

## 📁 Archivos del Módulo

```
c:\PROYECTOS\PRESUPUESTO\metales\
├── .github/workflows/
│   └── actualizar.yml          # 🤖 Automatización en la nube (GitHub Actions a las 6:00 PM)
├── index.html                  # 🌐 Interfaz Web Interactiva con Modo Administrador 1 Clic
├── data.js                     # 📊 Base de datos JavaScript exportada automáticamente
├── data.json                   # 📄 Base de datos JSON para APIs o consumo externo
├── actualizar_metales.py       # 🐍 Script principal de actualización en Python
├── actualizar_metales.bat      # ⚡ Ejecutable automático en 1 Clic (Python + Publicar en GitHub)
├── Precios_Metales_COP.xlsx    # 📈 Libro de Excel con Dashboard y datos diarios
└── README.md                   # 📝 Documentación del módulo
```

---

## 📐 Fórmula de Conversión a Pesos Colombianos por Gramo

$$\text{Precio Oro (COP/g)} = \left( \frac{\text{Precio Oro USD/oz}}{31.1034768} \right) \times \text{TRM (USD/COP)}$$

$$\text{Precio Plata (COP/g)} = \left( \frac{\text{Precio Plata USD/oz}}{31.1034768} \right) \times \text{TRM (USD/COP)}$$
