# Tarjetitas App 💳📊

¡Bienvenido! **Tarjetitas App** es una aplicación integral de gestión y optimización financiera orientada al control detallado de tarjetas de crédito. Desarrollada en Python utilizando el framework **Flet**, la aplicación resuelve un problema crítico del día a día: la descentralización de los resúmenes bancarios y la complejidad de calcular cuotas futuras, intereses y fechas de cierre.

Esta herramienta permite a los usuarios proyectar sus gastos reales mes a mes, asegurando una planificación financiera saludable y automatizada.

---

## 🚀 Características Principales

* **Dashboard General del Estado de Cuentas:** Vista unificada de múltiples tarjetas de crédito, destacando de forma clara los montos totales, fechas de cierre y fechas de vencimiento del período actual.
* **Proyección de Gastos Mensuales (Analytics):** Incorpora gráficos de barras dinámicos e interactivos que permiten visualizar la evolución del gasto y los compromisos financieros mes a mes.
* **Gestión Avanzada de Consumos:** Formulario inteligente para registrar gastos individuales especificando el comercio, la categoría, el monto inicial, el porcentaje de interés y el número de cuotas.
* **Cálculo Automatizado de Cuotas:** Distribución lógica en el tiempo de los consumos diferidos, impactando automáticamente en los meses correspondientes según las fechas de corte de cada tarjeta.
* **Historial Detallado e Interactivo:** Lista interactiva de los consumos del mes en curso con funciones integradas para modificar parámetros o eliminar registros de manera directa (operaciones CRUD completas).

---

## 🛠️ Tecnologías Utilizadas

* **Backend / Lógica de Negocio:** Python 🐍 (Procesamiento de datos, cálculo de cuotas e intereses cronológicos).
* **Frontend / Interfaz Gráfica:** Flet 💻 (Construcción de una UI/UX moderna, reactiva y totalmente adaptada a escritorio).
* **Análisis Visual:** Componentes gráficos nativos para la representación de métricas y barras de consumo.

---

## 📦 Estructura del Proyecto

* `main.py`: Punto de entrada y orquestador principal de la interfaz gráfica y navegación de la app.
* `assets/`: Recursos estáticos y multimedia utilizados para personalizar y estilizar la experiencia visual.
* `requirements.txt`: Archivo de configuración que centraliza las dependencias de producción para el entorno.
* `.gitignore` & `.vscode/`: Archivos de entorno para garantizar que el código se mantenga limpio y libre de archivos basura del sistema.

---

## ⚙️ Instalación y Ejecución Local

Si querés probar la aplicación de forma local en tu computadora, seguí estos pasos:

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/AlejoMartinez1995/tarjetitas-app.git](https://github.com/AlejoMartinez1995/tarjetitas-app.git)
   cd tarjetitas-app
   
### 📸 Vista previa de la aplicación

<p align="center">
  <img src="screenshots/dashboard.png" alt="Dashboard Principal" width="45%">
</p>
