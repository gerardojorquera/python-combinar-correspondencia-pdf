# 🎓 Generador Masivo de Certificados de Informática (Streamlit & ReportLab)

Esta aplicación web permite automatizar la creación masiva de diplomas y certificados académicos para el área de tecnología a partir de una planilla de datos Excel (.xlsx). Cuenta con soporte para añadir logos corporativos personalizados y exportar los resultados en formato individual comprimido o unificado.

---

## ✨ Funcionalidades Principales

*   **Validación Elástica de Columnas:** Interfaz inteligente que mapea los campos obligatorios del archivo Excel ignorando de forma automática el uso de mayúsculas, minúsculas, espacios o tildes.
*   **Campos de Datos Simplificados:** Estructura de cabeceras optimizada mediante identificadores cortos (`tokens-slug`) para facilitar la preparación de los archivos de origen.
*   **Identidad Corporativa:** Selector dinámico en el panel lateral para cargar el logo de la empresa que dicta el curso (top izquierdo) y el de la empresa que lo recibe (top derecho).
*   **Doble Modalidad de Salida:** 
    *   Generación de archivos PDF empaquetados uno a uno dentro de un comprimido `.zip`.
    *   Consolidación de toda la nómina en un único documento de impresión PDF multi-página.
*   **Monitoreo en Tiempo Real:** Barra de progreso e indicador de estado interactivo que detalla el número del certificado actual en proceso de diseño junto al nombre del participante.
*   **Diseño Ejecutivo Integrado:** Plantilla con orientación horizontal estándar, marcos decorativos institucionales en azul y oro, y control automático de desborde tipográfico para evitar firmas huérfanas.

---

## 📊 Estructura del Archivo Excel

Para que la aplicación procese los datos correctamente, tu archivo de Excel debe contener exactamente las siguientes columnas (el orden no importa y los nombres no discriminan tildes ni mayúsculas):

| Nombre de Columna | Descripción / Ejemplo |
| :--- | :--- |
| `titulo-certificado` | Nombre principal del diploma (Ej: *Certificado de Aprobación*) |
| `nombre-participante` | Nombre completo del alumno (Ej: *Francisco Reyes Retamal*) |
| `texto-central` | Glosa o descripción técnica del programa cursado e hitos cubiertos |
| `fecha-emision` | Glosa de localidad y tiempo (Ej: *Santiago de Chile - 13 de agosto de 2026*) |
| `nombre-relator` | Nombre completo del instructor o profesor |
| `empresa-relator` | Cargo y entidad del profesor (Ej: *Relator - NETCapacitaciones*) |
| `nombre-coordinador` | Nombre de la jefatura o ministro de fe del programa |
| `empresa-coordinador` | Cargo y empresa de coordinación (Ej: *Jefe de RRHH - MediSoft*) |

---

## 📋 Prerrequisitos del Sistema

Antes de iniciar la instalación local, asegúrate de tener instalado en tu computadora:

*   **Python 3.9 o superior**
*   **Gestor de paquetes `pip`**

---

## 🚀 Guía de Ejecución Local

Sigue paso a paso estas instrucciones en la consola de comandos de tu sistema para levantar el entorno de desarrollo:

### 1. Clonar o descargar los archivos del proyecto
Crea una carpeta en tu máquina e introduce el archivo ejecutable de la app (`app.py` o el nombre con el que guardaste el script de Streamlit) y el archivo de requerimientos (`requirements.txt`).

### 2. Crear y activar un entorno virtual (Recomendado)
Para asegurar el aislamiento de las librerías gráficas:

*   **En Windows:**
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```
*   **En macOS/Linux:**
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

### 3. Instalar las dependencias de Python
Instala las librerías encargadas de la manipulación de datos y renderizado vectorial de documentos PDF:
```bash
pip install -r requirements.txt
```

*(Nota: Asegúrate de que tu archivo `requirements.txt` contenga textualmente: `streamlit`, `pandas`, `openpyxl`, `reportlab`, `Pillow`, y `PyPDF2`).*

### 4. Iniciar la aplicación web
Lanza el servidor local a través de la terminal:
```bash
streamlit run app.py
```

Al terminar de procesar, se abrirá de manera automática una ventana en tu navegador web predeterminado apuntando a la dirección local del servicio, usualmente: `http://localhost:8501`.
