# TFG-Miguel-Triguero-Romero-Excel-VBA
Herramienta integral en Excel-VBA para la gestión de resultados académicos en el CUD


<br />

<div align="center">
  <h1>Sistema de Análisis Académico Automatizado</h1>
  <p><i>Desarrollo de una herramienta integral en Excel-VBA para la gestión de resultados y generación de informes estadísticos.</i></p>
  <p><b>Trabajo Fin de Grado - Ingeniería Mecánica (CUD / ENM)</b></p>
</div>

---

## Resumen del Proyecto
Esta herramienta ha sido diseñada para optimizar la gestión docente en el **Centro Universitario de la Defensa**. Mediante programación avanzada en **VBA**, la aplicación permite transformar actas de calificaciones en diagnósticos detallados, evaluando tanto el rendimiento individual del alumnado como la calidad psicométrica de las pruebas de evaluación.

---

## Funcionalidades Principales

### Inteligencia de Datos y Patrones
La herramienta no solo calcula notas, sino que utiliza algoritmos para identificar perfiles específicos:
* **Identificación de "Alumnos Tipo":** Clasificación automática en perfiles (Constante, Evolutivo Positivo, Riesgo Académico, etc.).
* **Caracterización de "Exámenes Tipo":** Análisis de la prueba (Discriminativa, Fácil, Difícil, etc.) basado en índices de dificultad y discriminación.

### Parámetros Estadísticos
Cálculo automatizado de métricas de alto nivel:
* **Estadística Descriptiva:** Media, Mediana, Desviación Típica.
* **Análisis de Distribución:** Coeficientes de Asimetría y Curtosis.
* **Métricas Educativas:** Índice de Dificultad e Índice de Discriminación.

### Generación de Informes y Visualización
* **Informes Narrativos:** Traducción de datos numéricos a texto coherente de forma automatizada.
* **Gráficos Dinámicos:** Generación síncrona de Gráficos.
* **Exportación:** Generación de informes en formato PDF.

---

## Vista Previa (Capturas)

| Interfaz de Usuario | Análisis Visual  | Informe Narrativo|
| :---: | :---: | :---: |
| <img src="img/image.png" width="300" /> | <img src="img/informe_texto.png" width="300" /> | <img src="img/graficos.png" width="300" /> |
| *Formularios de control VBA* | *Diagnóstico de Alumno Tipo* | *Distribución y Boxplots* |

---

## Estructura del Repositorio

* **`/src`**: Módulos de código fuente.
* **`/docs`**: Documentación del proyecto y Memoria del TFG.
* **`/test`**: Bases de datos utilizadas para la validación (generadas mediante IA).
* **`Herramienta_TFG_MMTT.xlsm`**: Archivo principal ejecutable.

---

## Instrucciones de Uso
1.  **Descargar:** Clona el repositorio o descarga el archivo `.xlsm`.
2.  **Seguridad:** Al abrir el archivo, haz clic en **"Habilitar macros"**.
3.  **Carga:** Introduce las notas en la hoja "Notas".
4.  **Ejecución:** Utiliza el panel de control para generar los informes deseados (Examen, Alumno o Acta Final).

---

## Autor y Dirección
* **Autor:** Miguel Triguero Romero
* **Directores:** Sergio Borrallo Tirado
* **Institución:** Centro Universitario de la Defensa en la Escuela Naval Militar (Universidade de Vigo).
* **Curso:** 2025-2026

---

<div align="center">
  <sub>Este proyecto ha sido desarrollado como parte del Grado en Ingeniería Mecánica.</sub>
</div>
