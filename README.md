# 💊 Sistema de Inventario de Medicamentos

<div align="center">

**Sistema profesional de gestión de inventario farmacéutico con interfaz gráfica moderna**

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-Apache%202.0-green.svg)](LICENSE)
[![PySide6](https://img.shields.io/badge/GUI-PySide6-orange.svg)](https://pypi.org/project/PySide6/)

[Características](#-características) •
[Instalación](#-instalación) •
[Uso](#-uso) •
[Documentación](#-documentación) •
[Licencia](#-licencia)

</div>

---

## 📋 Descripción

**Inventario de Medicamentos** es una aplicación de escritorio completa diseñada para gestionar el inventario de farmacias y establecimientos médicos. Combina una interfaz gráfica moderna y elegante (GUI) con una interfaz de línea de comandos (CLI) para máxima flexibilidad. Los datos se almacenan en archivos Excel para facilitar su edición y respaldo.

### 🎯 Propósito

Esta aplicación está diseñada para:
- **Farmacias pequeñas y medianas**: Control de stock en tiempo real
- **Consultorios médicos**: Gestión de medicamentos disponibles
- **Botiquines empresariales**: Control de suministros médicos
- **Uso personal**: Seguimiento de medicamentos en casa

---

## ✨ Características

### 🖥️ Interfaz Gráfica (GUI)

#### 📦 Gestión de Inventario
- ✅ **Búsqueda inteligente** por nombre de medicamento o componente activo
- ✅ **Autocompletado** con sugerencias en tiempo real
- ✅ **Visualización en tabla** con ordenamiento por columnas
- ✅ **Sistema de alias** para medicamentos (ej: "dolocordal" → "acetaminofén")
- ✅ **Actualización automática** cada 30 segundos
- ✅ **Creación de plantillas** Excel desde la aplicación
- ✅ **Refrescado manual** de datos desde Excel

#### 💰 Registro de Ventas
- ✅ **Interfaz intuitiva** para registro rápido de ventas
- ✅ **Validación de stock** en tiempo real
- ✅ **Cálculo automático** de totales
- ✅ **Alertas visuales** de stock bajo
- ✅ **Historial completo** de todas las transacciones

#### 📊 Estadísticas y Reportes
- ✅ **Gráficos visuales** de medicamentos más vendidos (Top 10)
- ✅ **Métricas en tiempo real**: total vendido, ganancia total
- ✅ **Alertas de stock bajo** (menos de 5 unidades)
- ✅ **Exportación automática** de estadísticas a Excel
- ✅ **Dashboard informativo** con tarjetas resumen

#### 🎨 Diseño Moderno
- ✅ **Tema oscuro profesional** para reducir fatiga visual
- ✅ **Notificaciones toast** no intrusivas
- ✅ **Animaciones suaves** entre pestañas
- ✅ **Responsive design** adaptable a diferentes resoluciones
- ✅ **Atajos de teclado** para acceso rápido (Ctrl+B, Ctrl+S, etc.)

### 💻 Interfaz de Línea de Comandos (CLI)

```bash
# Búsqueda de medicamentos
python inventario_medicamentos.py --search "aspirina"

# Registro de ventas
python inventario_medicamentos.py --sell "aspirina,10"

# Visualización de estadísticas
python inventario_medicamentos.py --stats
```

---

## 🔧 Requisitos del Sistema

### Software Necesario

| Requisito | Versión Mínima | Recomendada |
|-----------|----------------|-------------|
| **Python** | 3.8+ | 3.10+ |
| **Sistema Operativo** | Windows 10, macOS 10.14, Linux | Windows 11, macOS 12+, Ubuntu 22.04+ |
| **Memoria RAM** | 2 GB | 4 GB+ |
| **Espacio en disco** | 500 MB | 1 GB |

### Dependencias de Python

- **pandas** (≥1.3.0) - Manipulación y análisis de datos
- **PySide6** (≥6.0.0) - Framework de interfaz gráfica
- **matplotlib** (≥3.4.0) - Gráficos y visualizaciones
- **openpyxl** (≥3.0.0) - Lectura y escritura de archivos Excel

---

## 📥 Instalación

### Paso 1: Verificar Python

Asegúrate de tener Python 3.8 o superior instalado:

```bash
python --version
# o en algunos sistemas
python3 --version
```

Si no tienes Python instalado, descárgalo desde [python.org](https://www.python.org/downloads/).

### Paso 2: Clonar o Descargar el Repositorio

**Opción A: Clonar con Git**
```bash
git clone https://github.com/Eddym06/inventario_medicamentos.git
cd inventario_medicamentos
```

**Opción B: Descargar ZIP**
1. Descarga el ZIP desde [GitHub](https://github.com/Eddym06/inventario_medicamentos)
2. Extrae el contenido
3. Abre una terminal en la carpeta extraída

### Paso 3: Instalar Dependencias

Es **altamente recomendado** usar un entorno virtual:

```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno virtual
# En Windows:
venv\Scripts\activate
# En macOS/Linux:
source venv/bin/activate

# Instalar dependencias
pip install pandas PySide6 matplotlib openpyxl
```

**Instalación sin entorno virtual** (no recomendado):
```bash
pip install pandas PySide6 matplotlib openpyxl
```

### Paso 4: Verificar la Instalación

```bash
# Verificar que todas las dependencias estén instaladas
python -c "import pandas, PySide6, matplotlib, openpyxl; print('✅ Todas las dependencias instaladas correctamente')"
```

---

## 🚀 Uso

### Modo Interfaz Gráfica (GUI) - Recomendado

#### Iniciar la Aplicación

```bash
python inventario_medicamentos.py
```

La aplicación se abrirá con tres pestañas principales:

#### 📋 Pestaña "Inventario"
1. **Buscar medicamentos**: 
   - Escribe en la barra de búsqueda (mínimo 2 caracteres)
   - Usa el autocompletado para encontrar medicamentos rápidamente
   - Los resultados aparecen automáticamente en la tabla

2. **Crear plantilla Excel**:
   - Haz clic en "📄 Crear Plantilla" para generar un nuevo archivo Excel vacío
   - Esto reiniciará todos los contadores a cero

3. **Abrir Excel**:
   - Haz clic en "📂 Abrir Plantilla" para editar el archivo en Excel
   - Útil para agregar múltiples medicamentos manualmente

4. **Refrescar datos**:
   - Haz clic en "🔄 Refrescar Datos" para recargar cambios hechos en Excel

#### 💰 Pestaña "Ventas"
1. **Seleccionar medicamento**:
   - Escribe o selecciona el medicamento del autocompletado
   - El stock disponible aparecerá automáticamente

2. **Ajustar cantidad**:
   - Usa los botones +/- o escribe la cantidad directamente
   - El precio total se calcula automáticamente

3. **Registrar venta**:
   - Haz clic en "💰 Realizar Venta" o presiona `Ctrl+S`
   - La venta se guarda inmediatamente en el Excel

#### 📊 Pestaña "Estadísticas"
- Visualiza métricas en tiempo real
- Revisa el Top 10 de medicamentos más vendidos
- Consulta alertas de stock bajo
- Haz clic en "🔄 Actualizar Estadísticas" para refrescar manualmente

### Modo Línea de Comandos (CLI)

#### Buscar Medicamentos

```bash
# Búsqueda por nombre
python inventario_medicamentos.py --search "aspirina"

# Búsqueda por componente
python inventario_medicamentos.py --search "ácido acetilsalicílico"
```

#### Registrar Ventas

```bash
# Formato: --sell "nombre_medicamento,cantidad"
python inventario_medicamentos.py --sell "aspirina,10"
python inventario_medicamentos.py --sell "acetaminofén,5"
```

#### Ver Estadísticas

```bash
python inventario_medicamentos.py --stats
```

Salida ejemplo:
```
Total Vendidos: 150
Ganancia Total: $4,320.00
Top 10: {'aspirina': 45, 'acetaminofén': 38, ...}
Stock Bajo: [{'Medicamento': 'Ibuprofeno', 'Cantidad': 3}]
```

---

## 📁 Estructura del Archivo Excel

El sistema requiere un archivo **Inventario.xlsx** con la siguiente estructura:

### 🗂️ Hoja "Inventario"

| Medicamento | Componente | Cantidad | Precio Unitario |
|------------|------------|----------|-----------------|
| Aspirina | Ácido acetilsalicílico | 100 | 5.50 |
| Acetaminofén | Paracetamol | 50 | 3.25 |
| Ibuprofeno | Ibuprofeno | 75 | 7.00 |

**Columnas requeridas:**
- `Medicamento` (texto): Nombre del medicamento
- `Componente` (texto): Componente activo principal
- `Cantidad` (número): Stock disponible (≥0)
- `Precio Unitario` (número): Precio por unidad (≥0)

### 🗂️ Hoja "Ventas"

| Fecha | Medicamento | Cantidad | Precio Total |
|-------|------------|----------|--------------|
| 2025-11-13 10:30:00 | Aspirina | 5 | 27.50 |
| 2025-11-13 11:45:00 | Acetaminofén | 2 | 6.50 |

**Columnas automáticas:**
- `Fecha` (texto): Timestamp automático (YYYY-MM-DD HH:MM:SS)
- `Medicamento` (texto): Nombre del medicamento vendido
- `Cantidad` (número): Unidades vendidas
- `Precio Total` (número): Cantidad × Precio Unitario

### 🗂️ Hoja "Estadisticas"

Generada automáticamente con:
- Total de unidades vendidas
- Ganancia total acumulada
- Top 10 medicamentos más vendidos
- Alertas de stock bajo (menos de 5 unidades)

### 📝 Notas Importantes

- ⚠️ **No modifiques los nombres de las hojas** (deben ser exactamente "Inventario", "Ventas", "Estadisticas")
- ⚠️ **No cambies los nombres de las columnas** (respeta mayúsculas y tildes)
- ⚠️ **Usa números, no texto** para cantidades y precios
- ⚠️ **No uses separadores de miles** en números (usa `1000` no `1,000`)
- ✅ **Guarda como .xlsx** (formato moderno de Excel)

---

## 🔍 Características Avanzadas

### Sistema de Alias

Configura nombres alternativos para medicamentos comunes editando `alias.json`:

```json
{
  "dolocordal": "acetaminofén",
  "aspirina": "ácido acetilsalicílico",
  "advil": "ibuprofeno"
}
```

Después, buscar "dolocordal" encontrará automáticamente "acetaminofén".

### Atajos de Teclado

| Atajo | Acción |
|-------|--------|
| `Ctrl+B` | Buscar en inventario |
| `Ctrl+L` | Limpiar búsqueda |
| `Ctrl+S` | Realizar venta |
| `?` | Abrir ayuda (clic en botón) |

### Auto-actualización

La aplicación verifica cambios en el Excel cada 30 segundos y recarga automáticamente los datos.

---

## 🛠️ Solución de Problemas

### ❌ "ModuleNotFoundError: No module named 'pandas'"

**Solución**: Instala las dependencias
```bash
pip install pandas PySide6 matplotlib openpyxl
```

### ❌ "FileNotFoundError: Archivo 'Inventario.xlsx' no encontrado"

**Solución**: Crea la plantilla desde la aplicación
1. Abre la aplicación (se cerrará con error)
2. Crea manualmente el archivo Excel con la estructura indicada
3. O usa el botón "Crear Plantilla" en la GUI

### ❌ "Columnas requeridas faltantes"

**Solución**: Verifica que la hoja "Inventario" tenga exactamente estas columnas:
- Medicamento
- Componente
- Cantidad
- Precio Unitario

### ❌ La aplicación no guarda los datos

**Solución**:
- Cierra Excel si está abierto
- Verifica permisos de escritura en la carpeta
- Desactiva "Solo lectura" en las propiedades del archivo

### ❌ Los gráficos no aparecen

**Solución**:
```bash
pip install --upgrade matplotlib
```

### 🆘 Centro de Ayuda Integrado

Dentro de la aplicación, haz clic en el botón **?** (esquina inferior izquierda) para acceder a:
- Guía de uso completa
- Solución de problemas paso a paso
- Mejores prácticas
- Preguntas frecuentes

---

## 📚 Documentación Adicional

### Estructura del Código

```
inventario_medicamentos/
│
├── inventario_medicamentos.py  # Aplicación principal
├── Inventario.xlsx            # Base de datos (generado)
├── alias.json                 # Alias de medicamentos (opcional)
├── README.md                  # Esta documentación
└── LICENSE                    # Licencia Apache 2.0
```

### Clases Principales

- **`DataManager`**: Gestiona carga, guardado y cache de datos
- **`InventarioTab`**: Interfaz de búsqueda e inventario
- **`VentaTab`**: Interfaz de registro de ventas
- **`EstadisticasTab`**: Interfaz de estadísticas y reportes
- **`MainWindow`**: Ventana principal con sistema de notificaciones

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Para cambios importantes:

1. **Fork** el repositorio
2. **Crea una rama** para tu feature (`git checkout -b feature/AmazingFeature`)
3. **Commit** tus cambios (`git commit -m 'Add: amazing feature'`)
4. **Push** a la rama (`git push origin feature/AmazingFeature`)
5. **Abre un Pull Request**

### Áreas de Mejora Sugeridas
- 🔐 Sistema de autenticación de usuarios
- ☁️ Sincronización en la nube
- 📱 Versión móvil con React Native
- 🔔 Notificaciones por correo de stock bajo
- 📊 Reportes avanzados en PDF
- 🌐 Soporte multi-idioma

---

## 📄 Licencia

Este proyecto está licenciado bajo la **Apache License 2.0**. Consulta el archivo [LICENSE](LICENSE) para más detalles.

### Resumen de la Licencia

✅ **Permitido**:
- Uso comercial
- Modificación
- Distribución
- Uso privado

⚠️ **Requerido**:
- Incluir aviso de licencia y copyright
- Documentar cambios realizados

❌ **Prohibido**:
- Uso de marcas registradas
- Responsabilidad del autor

---

## 📞 Soporte y Contacto

### 🐛 Reportar Problemas

Si encuentras un bug o tienes una sugerencia:
1. Revisa primero los [Issues existentes](https://github.com/Eddym06/inventario_medicamentos/issues)
2. Si es nuevo, [crea un Issue](https://github.com/Eddym06/inventario_medicamentos/issues/new) con:
   - Descripción clara del problema
   - Pasos para reproducirlo
   - Versión de Python y sistema operativo
   - Capturas de pantalla (si aplica)

### 💬 Discusiones

Para preguntas generales o ideas, usa la sección [Discussions](https://github.com/Eddym06/inventario_medicamentos/discussions).

---

## 🌟 Agradecimientos

- **PySide6**: Framework de interfaz gráfica
- **Pandas**: Análisis de datos
- **Matplotlib**: Visualización de gráficos
- **OpenPyXL**: Manejo de archivos Excel

---

## 📊 Estado del Proyecto

🟢 **Activo**: En desarrollo y mantenimiento continuo.

### Versión Actual: 1.0.0 (Pro)

**Últimas actualizaciones**:
- ✅ Interfaz gráfica completa con tema oscuro
- ✅ Sistema de notificaciones toast
- ✅ Auto-actualización de datos
- ✅ Gráficos estadísticos interactivos
- ✅ Centro de ayuda integrado
- ✅ Soporte CLI y GUI

---

<div align="center">

**Hecho con ❤️ para la gestión eficiente de inventarios farmacéuticos**

⭐ Si este proyecto te ayuda, considera darle una estrella en GitHub

[⬆ Volver arriba](#-sistema-de-inventario-de-medicamentos)

</div>
