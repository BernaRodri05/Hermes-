═══════════════════════════════════════════════════════════════════
  HERMES - VERSIÓN INTEGRADA CON PROCESADOR DE EXCEL
═══════════════════════════════════════════════════════════════════

AUTOR: Berna - 2025
VERSIÓN: Integrada con Procesador de Excel

═══════════════════════════════════════════════════════════════════
  NOVEDADES DE ESTA VERSIÓN
═══════════════════════════════════════════════════════════════════

Esta versión integra el procesador de Excel directamente en Hermes.
Ahora puedes:

✓ Cargar archivos Excel con datos de contactos
✓ Seleccionar qué columnas incluir en el mensaje
✓ Crear mensajes personalizados con plantillas
✓ Generar automáticamente URLs de WhatsApp
✓ Exportar un Excel con solo los URLs en la primera columna
✓ Enviar mensajes masivos directamente desde Hermes

═══════════════════════════════════════════════════════════════════
  CÓMO FUNCIONA
═══════════════════════════════════════════════════════════════════

PASO 1: DETECTAR DISPOSITIVOS
   - Conecta tu celular por USB
   - Habilita depuración USB en el celular
   - Haz clic en "Detectar Dispositivos"

PASO 2: CARGAR Y PROCESAR EXCEL
   - Haz clic en "Cargar y Procesar Excel"
   - Selecciona tu archivo Excel
   - El archivo DEBE tener al menos una columna con "Telefono" en el nombre
   
   En la ventana de configuración:
   - Selecciona las columnas que quieres incluir en el mensaje
   - Escribe una plantilla de mensaje (opcional)
   - Usa {NombreColumna} para insertar valores
   - Ejemplo: "Hola {Nombre}, te contactamos de {Empresa}"
   
   - Haz clic en "Procesar y Generar URLs"
   - Guarda el Excel procesado (solo tendrá URLs en la columna A)

PASO 3: INICIAR ENVÍO
   - Haz clic en "INICIAR ENVÍO"
   - Los mensajes se enviarán automáticamente

═══════════════════════════════════════════════════════════════════
  FORMATO DEL EXCEL DE ENTRADA
═══════════════════════════════════════════════════════════════════

Tu Excel debe tener:

1. Al menos UNA columna con "Telefono" en el nombre
   Ejemplos: Telefono, Telefono1, Telefono2, telefono_celular

2. Los números de teléfono deben estar SIN el prefijo +549
   Ejemplo: 1234567890 (no +5491234567890)

3. Si tienes varios números separados por guión, se procesarán todos
   Ejemplo: 1234567890-0987654321

4. Otras columnas pueden tener cualquier nombre
   Ejemplos: Nombre, Empresa, Dirección, etc.

═══════════════════════════════════════════════════════════════════
  EJEMPLO DE EXCEL DE ENTRADA
═══════════════════════════════════════════════════════════════════

Nombre          | Telefono1    | Telefono2    | Empresa
----------------|--------------|--------------|-------------
Juan Perez      | 1234567890   | 0987654321   | Empresa A
Maria Lopez     | 1122334455   | 5544332211   | Empresa B
Carlos Gomez    | 9988776655   |              | Empresa C

═══════════════════════════════════════════════════════════════════
  RESULTADO DEL PROCESAMIENTO
═══════════════════════════════════════════════════════════════════

El Excel procesado tendrá SOLO una columna con URLs:

URL
-----------------------------------------------------------------
https://wa.me/5491234567890?text=Hola%20Juan%20Perez...
https://wa.me/5490987654321?text=Hola%20Juan%20Perez...
https://wa.me/5491122334455?text=Hola%20Maria%20Lopez...
https://wa.me/5495544332211?text=Hola%20Maria%20Lopez...
https://wa.me/5499988776655?text=Hola%20Carlos%20Gomez...

Este Excel está listo para ser usado por Hermes para el envío masivo.

═══════════════════════════════════════════════════════════════════
  CARACTERÍSTICAS MANTENIDAS DE HERMES ORIGINAL
═══════════════════════════════════════════════════════════════════

✓ Mismo estilo visual y diseño
✓ Configuración de tiempos entre mensajes
✓ Pausar y reanudar envíos
✓ Cancelar envíos en curso
✓ Registro de actividad en tiempo real
✓ Estadísticas de envío
✓ Progreso visual
✓ Tiempo estimado de finalización

═══════════════════════════════════════════════════════════════════
  INSTALACIÓN
═══════════════════════════════════════════════════════════════════

1. Ejecuta INSTALAR.bat (solo la primera vez)
2. Ejecuta EJECUTAR.bat para iniciar Hermes

═══════════════════════════════════════════════════════════════════
  NOTAS IMPORTANTES
═══════════════════════════════════════════════════════════════════

• Los números de teléfono se procesan automáticamente con el prefijo 549
• Si una fila no tiene número de teléfono, se omite automáticamente
• Puedes tener múltiples columnas de teléfono (Telefono1, Telefono2, etc.)
• Los mensajes se personalizan usando {NombreColumna} en la plantilla
• El Excel procesado tiene SOLO los URLs, sin datos personales

═══════════════════════════════════════════════════════════════════
  SOPORTE
═══════════════════════════════════════════════════════════════════

Para problemas o consultas, contacta a Berna.

═══════════════════════════════════════════════════════════════════

