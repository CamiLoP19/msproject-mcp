# Microsoft Project MCP Server (`msproject-mcp`)

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6.svg)](https://www.microsoft.com/windows)
[![Protocol: Model Context Protocol](https://img.shields.io/badge/protocol-MCP-8A2BE2.svg)](https://modelcontextprotocol.io)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

Servidor **Model Context Protocol (MCP)** para controlar, automatizar y consultar **Microsoft Project de escritorio (archivos `.mpp` y `.xml`)** en tiempo real mediante asistentes de inteligencia artificial (Claude, Cursor, Antigravity, VS Code, etc.).

A diferencia de lectores estáticos de archivos, este servidor interactúa directamente con la aplicación de escritorio de MS Project mediante la API COM nativa de Windows (`win32com`), permitiendo ver los cambios en pantalla en vivo, recalcular cronogramas, asignar recursos y exportar resultados.

---

## 🌟 Características Principales

* 🔄 **Control Total de Proyectos:** Crear, abrir, activar, guardar y cerrar archivos `.mpp` y `.xml`.
* 📋 **Gestión de Tareas y EDT / WBS:** Creación de tareas con duraciones, fechas, sangrías jerárquicas (`indent_task`), hitos y vinculación de dependencias (`predecessors`).
* 👥 **Gestión de Recursos y Costos:** Alta y asignación de recursos de tipo **Trabajo**, **Material** y **Costo**, con tarifas por hora y asignación de unidades.
* 📈 **Ruta Crítica y Línea Base:** Fijación de líneas base (`BaselineSave`), cálculo instantáneo de la ruta crítica y reportes de estado con desviaciones y tareas atrasadas.
* 🌐 **Soporte Bilingüe Automático:** Compatible tanto con Microsoft Project en **español** como en **inglés** (normaliza automáticamente duraciones como `1w` a `1s` y predecesoras como `FS` a `FC`).
* 📄 **Exportación a PDF:** Generación directa de reportes en PDF del cronograma.

---

## 📋 Requisitos del Sistema

Para ejecutar este servidor MCP se requiere:
1. **Sistema Operativo:** Windows 10 o Windows 11.
2. **Microsoft Project:** Instalación de escritorio de MS Project (versiones 2016, 2019, 2021 o Microsoft 365 Project).
3. **Python:** Versión 3.10 o superior (recomendado instalar con [uv](https://github.com/astral-sh/uv) o desde python.org).

---

## 🚀 Instalación y Puesta en Marcha

### Opción 1: Con `uv` (Recomendado - Rápido)

1. **Clonar el repositorio:**
   ```powershell
   git clone https://github.com/tu-usuario/msproject-mcp.git
   cd msproject-mcp
   ```

2. **Crear el entorno e instalar dependencias:**
   ```powershell
   uv venv
   .venv\Scripts\activate
   uv sync
   ```

### Opción 2: Con `pip` y Python estándar

1. **Clonar el repositorio:**
   ```powershell
   git clone https://github.com/tu-usuario/msproject-mcp.git
   cd msproject-mcp
   ```

2. **Crear y activar el entorno virtual:**
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```powershell
   pip install -e .
   ```

---

## ⚙️ Configuración en Clientes de IA

Copia la configuración correspondiente según la herramienta que utilices.

### 1. Claude Desktop
Abre tu archivo de configuración en `%APPDATA%\Claude\claude_desktop_config.json` y añade:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\ruta\\a\\tu\\msproject-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "msproject_mcp"],
      "env": {
        "MSPROJECT_VISIBLE": "1"
      }
    }
  }
}
```

### 2. Cursor
En Cursor, abre **Settings > Features > MCP** o edita el archivo `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\ruta\\a\\tu\\msproject-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "msproject_mcp"],
      "env": {
        "MSPROJECT_VISIBLE": "1"
      }
    }
  }
}
```

### 3. VS Code (Cline / Roo Code / Continue)
En los ajustes de extensiones MCP de VS Code:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\ruta\\a\\tu\\msproject-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "msproject_mcp"],
      "env": {
        "MSPROJECT_VISIBLE": "1"
      }
    }
  }
}
```

### 4. Antigravity IDE
Crea o edita el archivo `.agents/mcp_config.json` en la raíz de tu espacio de trabajo:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\ruta\\a\\tu\\msproject-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "msproject_mcp"],
      "env": {
        "MSPROJECT_VISIBLE": "1"
      }
    }
  }
}
```

> **Nota:** Reemplaza `C:\\ruta\\a\\tu\\msproject-mcp` con la ruta absoluta donde clonaste el repositorio en tu PC.

---

## 🛠️ Herramientas Disponibles (25 Herramientas)

### 📂 Proyectos
| Herramienta | Descripción |
| :--- | :--- |
| `create_project` | Crea un proyecto nuevo con fecha de inicio, título y guardado opcional. |
| `open_project` | Abre un archivo `.mpp` o `.xml` (con opción de solo lectura). |
| `list_open_projects` | Lista todos los proyectos abiertos actualmente en MS Project. |
| `switch_project` | Cambia el proyecto activo por nombre o ruta. |
| `save_project` | Guarda el proyecto activo o hace 'Guardar como' (.mpp o .xml). |
| `export_pdf` | Exporta el proyecto activo a un documento PDF. |
| `close_project` | Cierra el proyecto activo guardando o descartando cambios. |
| `get_project_info` | Devuelve metadatos, fechas de inicio/fin, conteo de tareas y recursos. |
| `update_project` | Modifica título, fecha de estado (`status_date`) o fecha de inicio. |

### 📝 Tareas y Esquema (EDT / WBS)
| Herramienta | Descripción |
| :--- | :--- |
| `list_tasks` | Lista las tareas del proyecto con filtro de nivel de esquema (`outline_level`). |
| `get_task` | Obtiene el detalle completo de una tarea (WBS, holgura, costos, variaciones). |
| `add_task` | Agrega una tarea indicando duración, predecesoras, nivel de esquema o hito. |
| `update_task` | Modifica nombre, fechas, % completado, duraciones, notas o predecesoras. |
| `indent_task` | Ajusta la sangría de una tarea (sangría a la derecha o izquierda para WBS). |
| `delete_task` | Elimina una tarea por su número de fila/ID. |

### 👤 Recursos y Asignaciones
| Herramienta | Descripción |
| :--- | :--- |
| `list_resources` | Lista los recursos dados de alta en el proyecto. |
| `add_resource` | Da de alta un recurso de tipo trabajo, material o costo con tarifa estándar. |
| `update_resource` | Modifica tarifas, unidades máximas o datos del recurso. |
| `delete_resource` | Elimina un recurso por nombre o ID. |
| `assign_resource` | Asigna un recurso a una tarea con porcentaje de dedicación o unidades. |
| `unassign_resource` | Desasigna un recurso de una tarea. |
| `list_assignments` | Muestra el mapeo tarea ↔ recurso con avance y horas de trabajo. |

### 📊 Seguimiento y Control
| Herramienta | Descripción |
| :--- | :--- |
| `set_baseline` | Guarda la Línea Base (`Baseline`) de todo el proyecto activo. |
| `critical_path` | Devuelve las tareas que componen la ruta crítica del proyecto. |
| `status_report` | Genera un reporte de avance: tareas en curso, atrasadas y desviaciones. |

---

## 🧪 Pruebas y Verificación

El repositorio incluye un conjunto de pruebas para validar la conexión COM con MS Project:

```powershell
# Ejecutar prueba de extremo a extremo:
python tests\smoke_test.py
```

Esta prueba abrirá Microsoft Project, creará un proyecto de prueba, estructurará tareas jerárquicas, asignará recursos, guardará línea base, calculará la ruta crítica y guardará el archivo `.mpp`.

---

## 🔧 Variables de Entorno

* `MSPROJECT_VISIBLE`: `1` (por defecto) para mantener la ventana de Project visible en pantalla, o `0` para ejecutarlo en segundo plano.
* `MSPROJECT_TIMEOUT`: Tiempo límite en segundos para llamadas COM (por defecto `90`).

---

## 📄 Licencia

Este proyecto se distribuye bajo la licencia **MIT**. Consulta el archivo `LICENSE` para más detalles.
