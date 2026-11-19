# VeriDocs - Sistema de Gestión y Renombrado Automático de Escaneos MES

Aplicación de grado industrial desarrollada para automatizar la recepción, reconocimiento OCR nativo de Windows, troceado por documentos, renombrado inteligente y replicación en red de los dossiers PDF escaneados en el puesto de verificación de taller. Sincroniza datos en tiempo real entre la Orden de Fabricación (OF), el catálogo maestro `codigosOCR.xlsx`, el Excel maestro de pedidos en curso de Aktis y los repositorios de Oficina Técnica.

---

## 🌟 Flujo de Trabajo Automatizado (Batch Continuo)

1. **Escaneo continuo sin clasificación manual:** El operario introduce todo el taco de hojas de la OF en el alimentador del escáner generando un único PDF.
2. **Detección Automática de OF:** VeriDocs analiza la cabecera (primeros 30%) con OCR nativo (`winocr` sobre Windows Media OCR) reconociendo formatos como `AYC/MO/00560`, `CESA/MO/00154`, `LOESS/MO/00087`, `MES/MO/05310`, `MO/00524`.
3. **Cruce en Vivo con Aktis:** Lee `LISTADO PEDIDOS EN CURSO AKTIS.xlsm` y recupera:
   * **Pedido ODOO** (Columna E)
   * **Descripción** (Columna I)
   * **Cliente** y datos de fabricación
4. **Troceado, Agrupación y Clasificación Inteligente:**
   * **Agrupación multicara:** Las páginas continuas con el mismo código de documento se consolidan en un único PDF.
   * **Gestión Dúplex Segura:** Filtro de páginas en blanco a 72 DPI en <15ms. Si el escáner genera reversos en blanco, se descartan automáticamente sin penalizar el rendimiento. Se preservan siempre los reversos con información útil y los documentos de pruebas FAT (`DC-0034` y `DC-0035`). El modo dúplex se puede activar/desactivar desde la interfaz.
   * **Hojas de Incidencia (`DC-0020`, `DC-0022`):** Si hay múltiples páginas, se indexan individualmente (`1-3`, `2-3`, `3-3`). Las hojas `DC-0022` se replican automáticamente en la carpeta general de incidencias.
   * **Páginas No Reconocidas:** Se aíslan en `REVISAR_PAGINAS_NO_IDENTIFICADAS - [OF].pdf` y se genera una alerta visual y sonora para el operario.
5. **Replicación Automática en Oficina Técnica:**
   * Si está habilitada la opción, busca la carpeta del proyecto en `\\172.30.0.10\Oficina_tecnica` (año prioritario -> código ODOO) y copia los documentos clasificados de forma transparente.

---

## 🚀 Mejoras de Arquitectura y Rendimiento Recientes

- **Optimización OCR x4:** Comprobación ultrarrápida de hoja en blanco antes de ejecutar OCR. Reutilización de bucle de eventos asíncronos y precálculo de patrones regex (`re.compile`) con tolerancia a separadores.
- **Renderizado de cabecera en una sola pasada:** Las rotaciones a 90° y 270° se aplican en memoria sobre la imagen rasterizada inicial, eliminando re-renderizados pesados de PDF.
- **Cola Protegida y Bloqueo de Concurrencia:** La sesión de escaneo procesa archivos secuencialmente con `threading.Lock()` evitando carreras entre la escucha automática, la subida manual de escritorio y la subida web.
- **IDs Monotónicos:** Los eventos de log y el historial de lotes se indexan con identificadores estrictamente crecientes, solucionando el congelamiento de logs tras 35 eventos.
- **Guardado Atómico de Configuración:** Escritura en fichero temporal + `os.replace` para blindar `config.json` frente a cortes de suministro eléctrico en PCs de taller.
- **Soporte Rutas Largas Windows (MAX_PATH):** Prefijo automático `\\?\` y `\\?\UNC\` para carpetas de red con anidación profunda.
- **Recarga en Caliente de Catálogo:** Actualizaciones a `codigosOCR.xlsx` se detectan y recargan dinámicamente sin reiniciar la aplicación.

---

## 🖥️ Modalidades de Uso

### 1. Interfaz de Escritorio (Desktop GUI)
* **Ejecutable nativo:**
  ```
  dist/VeriDocs/VeriDocs.exe
  ```
* **O arranque mediante script:** `iniciar_escritorio.bat`
* **Novedades de la GUI:**
  * Indicadores de estado de red (conexión con Oficina Técnica y escaneos pendientes en cola).
  * Selector visual de directorios y archivos mediante botones `...` en la ventana de configuración.
  * Selector de activación de **Modo Dúplex**.
  * Navegación secuencial de alertas activas (`[1/N]`, "Siguiente ▶", "Descartar").
  * Botones de acción directa: "📂 Abrir Carpeta OF", "📋 Copiar Ruta" y "📁 Abrir en OT" compatibles con lotes históricos.
  * Búsqueda en Aktis con debounce para evitar congelamientos al teclear.
  * Notificaciones Toast nativas de Windows (`winrt`) con fallback seguro.

### 2. Interfaz Web (Web Dashboard)
* **Arranque:** `iniciar_web.bat`
* **Acceso desde cualquier equipo o tablet en red local:**
  ```
  http://<IP-DEL-EQUIPO>:5055
  ```
* **Novedades Web:**
  * Zona Drag & Drop para arrastrar y procesar PDFs manualmente con barra de progreso en vivo.
  * Previsualización y descarga directa de PDFs generados desde el navegador (`/api/files/download`).
  * Indicadores pill de estado de Oficina Técnica y escaneos pendientes en cabecera.
  * Buscador interactivo en vivo del catálogo oficial de códigos OCR.
  * Sistema de notificaciones no bloqueantes tipo Toast.
  * Despliegue de historial de lotes con acceso a archivos y rutas.
  * Configuración completa desde modal web con protecciones de seguridad (límite 100MB, sanitización de rutas).

---

## ⚙️ Estructura de Ficheros

```
VeriDocs/
├── app_gui.py              # Interfaz gráfica de escritorio (CustomTkinter)
├── app_web.py              # Servidor y API REST web local (Flask)
├── veridocs_core.py        # Núcleo de lógica de negocio, lotes y replicación
├── ocr_service.py          # Motor de renderizado OCR nativo y detección de patrones
├── config.py               # Gestión atómica de configuración
├── config.json             # Parámetros y rutas activas del sistema
├── codigosOCR.xlsx         # Catálogo maestro de tipos de documentos
├── build_exe.py            # Script automatizado de empaquetado con PyInstaller
├── VeriDocs.spec           # Especificación de PyInstaller
├── requirements.txt        # Dependencias de Python verificadas
├── tests/
│   └── test_veridocs.py    # Suite de pruebas unitarias automatizadas
├── templates/
│   └── index.html          # Vista HTML del dashboard web
└── static/
    ├── css/style.css       # Estilos responsivos con fuentes fallback para PCs aislados
    └── js/app.js           # Lógica frontend reactiva e incremental
```

---

## 🧪 Ejecución de Pruebas Unitarias

Para validar el motor de OCR, detección de OFs, rutas UNC y gestión de sesiones:
```powershell
python -m unittest tests/test_veridocs.py
```

---

## 📦 Compilación a Ejecutable Windows (.exe)

Para regenerar el ejecutable de distribución:
```powershell
python build_exe.py
```
El resultado se generará en la carpeta `dist/VeriDocs/`.
