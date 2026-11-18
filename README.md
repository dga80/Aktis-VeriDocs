# VeriDocs - Sistema de Gestión y Renombrado Automático de Escaneos MES

Aplicación desarrollada para automatizar la recepción, reconocimiento OCR, troceado por documentos y renombrado de los dossiers PDF escaneados en el puesto de verificación de taller en base a la Orden de Fabricación (OF), el catálogo maestro `codigosOCR.xlsx` y el archivo maestro de pedidos en curso de Aktis.

---

## 🌟 Nueva Evolución: Escaneo Continuo por Lote (Batch Automatizado)

Ya no es necesario que el operario busque la OF ni seleccione manualmente qué documentos va a escanear:
1. **Un solo escaneo continuo:** El operario introduce todo el taco de hojas de la OF en el escáner y genera un **único archivo PDF**.
2. **Detección Automática de la OF:** VeriDocs analiza la cabecera (primeros 30%) de la página 1 con OCR nativo de Windows y detecta automáticamente la OF (ej: `AYC/MO/00560`, `CESA/MO/00154`, `LOESS/MO/00087`, `MES/MO/05310`).
3. **Búsqueda en Aktis:** Conecta en tiempo real con `LISTADO PEDIDOS EN CURSO AKTIS.xlsm` y recupera:
   * **Pedido ODOO** (Columna E)
   * **Descripción** (Columna I)
   * **Cliente** y datos adicionales
4. **Troceado y Agrupación Inteligente:** Analiza el pie de página de cada hoja contra el catálogo maestro `codigosOCR.xlsx`:
   * Las páginas con el mismo código se **agrupan en un único PDF** (ej: Hoja de verificación de 3 páginas).
   * **Incidencias (`DC-0020`, `DC-0022`):** Si hay más de 1 página de incidencias, se generan documentos independientes individuales con la nomenclatura `1-3`, `2-3`, `3-3`. Se duplican en la carpeta general de `incidencias` únicamente las páginas con código `DC-0022` (las `DC-0020` no se copian a dicha carpeta).
   * **Páginas no reconocidas:** Si alguna hoja tiene el pie borroso o sin código, se guarda en `REVISAR_PAGINAS_NO_IDENTIFICADAS - [OF].pdf` y **la app emite una alerta destacada en pantalla**.
5. **Guardado y Archivado:** Crea la carpeta de la OF en la red de la empresa, guarda los PDFs clasificados y archiva una copia del PDF original escaneado.

---

## 🚀 Modalidades de Uso

### Opción 1: Aplicación Ejecutable de Escritorio (.exe)
* Ubicación:
  ```
  dist/VeriDocs/VeriDocs.exe
  ```
* Se puede iniciar con doble clic o mediante `iniciar_escritorio.bat`.
* Incluye monitor en directo, avisos sonoros, panel de alertas, visor de último lote y buscador de OFs.

### Opción 2: Modo Web / Red Empresa
* Iniciar el servidor mediante `iniciar_web.bat`.
* Accesible desde cualquier navegador de la empresa sin instalar nada:
  ```
  http://<IP-DEL-EQUIPO>:5055
  ```

---

## ⚙️ Configuración y Catálogo de Códigos

* **Archivo Excel Aktis:** `\\172.30.0.10\Oficina_tecnica\99-PEDIDOS EN CURSO AKTIS\LISTADO PEDIDOS EN CURSO AKTIS.xlsm`
* **Carpeta de Escaneos:** `\\Mes-verificacio\escaneos`
* **Carpeta Base de Destino:** `\\Mes-verificacio\escaneos`
* **Catálogo OCR:** `codigosOCR.xlsx` (con los 24 tipos de documentos del sistema).
* **Parámetros OCR:** Cabecera (30%), Pie de página (15%). Modificables desde el menú **⚙️ Configuración**.
