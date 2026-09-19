# ⚡ Limpiador de Datos de Vatímetro

Esta herramienta es una aplicación de escritorio diseñada para procesar, limpiar y analizar automáticamente los registros de datos (archivos `.csv`) exportados por un vatímetro.

El programa transforma los datos en bruto en un reporte de Excel (`.xlsx`) profesional, interactivo y fácil de analizar, generando estadísticas de voltaje, corriente y potencia, además de gráficas y cálculos dinámicos de energía.

## ✨ Características Principales

* **Interfaz Gráfica Sencilla:** Utiliza ventanas emergentes para guiar al usuario, seleccionar archivos y configurar parámetros sin necesidad de usar la consola.
* **Procesamiento por Lotes:** Permite seleccionar uno o varios archivos CSV al mismo tiempo y los procesa de forma secuencial.
* **Filtro Inteligente de Ruido:** Ofrece la opción de ignorar los valores `0` (y ruido microscópico < 0.0001) para que no alteren las estadísticas reales de Media, Mediana y Mínimo.
* **Resumen Estadístico Automático:** Calcula Valores Máximos, Mínimos, Medios y Medianos para las variables principales (Voltaje Urms, Corriente Irms, Potencia Activa P y Potencia Aparente S) de hasta 3 elementos/fases.
* **Clasificación de Potencia (Estándar 62368-1):** Colorea automáticamente los valores de potencia activa en el Excel según los límites normativos (PS1, PS2, PS3).
* **Gráficas Automáticas:** Genera curvas continuas de Potencia Activa frente al tiempo para cada elemento, listas para visualizar en su propia pestaña.
* **Calculadora Dinámica de Energía:** Incluye una pestaña interactiva donde el usuario puede introducir un "Umbral de Disparo" y el "Tiempo de Ensayo", y el Excel calculará automáticamente la energía disipada (Julios) mediante fórmulas integradas.

## 🚀 Cómo usar la aplicación

No necesitas instalar Python ni conocimientos de programación para usar esta herramienta, solo ejecutar la aplicación compilada.

1. Dirígete a la carpeta `dist` de este repositorio.
2. Descarga y haz doble clic sobre el archivo ejecutable (`.exe`).
3. Lee el mensaje de bienvenida con las instrucciones y haz clic en "Aceptar".
4. Se abrirá un explorador de archivos: selecciona uno o varios archivos `.csv` generados por el vatímetro.
5. El programa te preguntará si deseas filtrar/ignorar los valores '0'. Selecciona "Sí" o "No" según necesites.
6. ¡Listo! El programa procesará los datos y generará automáticamente un archivo Excel (`_potencia.xlsx`) en la misma carpeta donde tenías tus archivos CSV originales.

> [!WARNING]
> **Nota para Desarrolladores:**
> La herramienta está diseñada para usarse directamente desde el `.exe`. Sin embargo, si deseas descargar el código fuente (`.py`) para modificarlo o ejecutarlo desde tu entorno, necesitarás tener instalado **Python** y las siguientes dependencias:
> ```bash
> pip install pandas numpy openpyxl
> ```
> *(Nota: `tkinter` y `os` ya vienen incluidos en la biblioteca estándar de Python).*
