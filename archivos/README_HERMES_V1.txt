═══════════════════════════════════════════════════════════════════
  HERMES V1 - ENVÍO MASIVO DE WHATSAPP
═══════════════════════════════════════════════════════════════════

AUTOR: Berna - 2025
VERSIÓN: 1.0 Final

═══════════════════════════════════════════════════════════════════
  CARACTERÍSTICAS
═══════════════════════════════════════════════════════════════════

✓ Procesador de Excel/CSV integrado
✓ Interfaz unificada con colores y tipografía Segoe UI
✓ Desplegables mejorados para mejor organización
✓ Selección precisa de columnas de teléfono
✓ Botones dinámicos para insertar campos
✓ Soporte completo para CSV con codificación latin-1
✓ Mensajes personalizados con plantillas
✓ Generación automática de URLs de WhatsApp
✓ Control total del proceso de envío

═══════════════════════════════════════════════════════════════════
  INSTALACIÓN
═══════════════════════════════════════════════════════════════════

1. Ejecuta INSTALAR.bat (solo la primera vez)
2. Espera a que se instalen las dependencias
3. Ejecuta EJECUTAR.bat para iniciar HERMES V1

═══════════════════════════════════════════════════════════════════
  CÓMO USAR
═══════════════════════════════════════════════════════════════════

PASO 1: DETECTAR DISPOSITIVOS
   - Conecta tu celular por USB
   - Habilita depuración USB
   - Haz clic en "🔍 Detectar Dispositivos"

PASO 2: CARGAR Y PROCESAR EXCEL/CSV
   - Haz clic en "📄 Cargar y Procesar Excel"
   - Selecciona tu archivo (.xlsx, .xls o .csv)
   
   En la ventana de configuración:
   
   📊 Paso 1: Información del Archivo (siempre visible)
      - Muestra nombre del archivo
      - Total de filas
      - Columnas de teléfono detectadas
   
   📞 Paso 2: Seleccionar Columnas de Teléfono (desplegable)
      - Haz clic para desplegar
      - Solo Telefono_1 está marcado por defecto
      - Marca las columnas de teléfono que quieras usar
      - Haz clic nuevamente para colapsar
   
   ✏️ Paso 3: Seleccionar Columnas para el Mensaje (desplegable)
      - Haz clic para desplegar
      - Marca las columnas que quieres incluir en el mensaje
      - Ejemplo: Razon Social, Cartera, Empresa
      - Haz clic nuevamente para colapsar
   
   💬 Paso 4: Plantilla de Mensaje (desplegable)
      - Haz clic para desplegar
      - Verás botones solo de las columnas marcadas en Paso 3
      - Haz clic en los botones para insertar {NombreColumna}
      - Escribe tu mensaje personalizado
      - Ejemplo: "Hola {Razon Social}, recordatorio de {Cartera}"
   
   - Haz clic en "✓ Procesar y Generar URLs"
   - Guarda el Excel procesado (solo URLs)

PASO 3: INICIAR ENVÍO
   - Haz clic en "▶ INICIAR ENVÍO"
   - Los mensajes se enviarán automáticamente
   - Puedes pausar o cancelar en cualquier momento

═══════════════════════════════════════════════════════════════════
  FORMATO DEL ARCHIVO
═══════════════════════════════════════════════════════════════════

EXCEL (.xlsx, .xls):
   - Debe tener al menos una columna con "telefono" en el nombre
   - Ejemplos: Telefono, Telefono_1, Telefono_2, telefono_celular
   - Los números deben estar SIN el prefijo +549
   - Ejemplo: 1161598799 (no +5491161598799)

CSV (.csv):
   - Mismo formato que Excel
   - Soporta codificación: latin-1, utf-8, cp1252
   - Delimitadores: ; , \t |
   - Detección automática de codificación y delimitador

MÚLTIPLES NÚMEROS:
   - Puedes separar números con guión: 1161598799-1162570366
   - Se procesarán ambos números automáticamente

═══════════════════════════════════════════════════════════════════
  COLORES Y DISEÑO
═══════════════════════════════════════════════════════════════════

HERMES V1 usa una paleta de colores unificada:

🔵 Azul (#4285F4)   - Acciones principales, header
🟢 Verde (#1DB954)  - Éxito, confirmaciones
🟠 Naranja (#FDB913) - Progreso, advertencias
⚪ Gris (#f8f9fa)   - Fondo general

Tipografía: Segoe UI en toda la aplicación

═══════════════════════════════════════════════════════════════════
  DESPLEGABLES MEJORADOS
═══════════════════════════════════════════════════════════════════

Los pasos 2, 3 y 4 son desplegables:

▼ Colapsado (cerrado)
   - Solo se ve el encabezado
   - Haz clic para desplegar

▲ Desplegado (abierto)
   - Se muestra el contenido completo
   - Haz clic para colapsar

Beneficios:
   - Interfaz más limpia
   - Menos scroll
   - Enfoque en un paso a la vez

═══════════════════════════════════════════════════════════════════
  BOTONES DINÁMICOS
═══════════════════════════════════════════════════════════════════

En el Paso 4, los botones se actualizan automáticamente:

- Si marcas: Nombre, Empresa
  → Ves botones: [Nombre] [Empresa]

- Si agregas: Ciudad
  → Aparece: [Nombre] [Empresa] [Ciudad]

- Si desmarcar Empresa
  → Quedan: [Nombre] [Ciudad]

Haz clic en un botón para insertar {NombreColumna} en el mensaje.

═══════════════════════════════════════════════════════════════════
  EJEMPLO COMPLETO
═══════════════════════════════════════════════════════════════════

ARCHIVO: InformedeCuentas.csv

Columnas:
- Razon Social
- Telefono_1
- Telefono_2
- Cartera
- Sueldo

PASO 1: Información
   📊 InformedeCuentas.csv
   📝 Total: 100 filas
   📞 Telefono_1, Telefono_2

PASO 2: Teléfonos
   ☑ Telefono_1 (marcado)
   ☐ Telefono_2

PASO 3: Columnas
   ☑ Razon Social
   ☑ Cartera
   ☐ Sueldo

PASO 4: Mensaje
   Botones: [Razon Social] [Cartera]
   
   Mensaje:
   "Hola {Razon Social}, recordatorio de pago de {Cartera}"

RESULTADO:
   URLs generados: 100
   Solo con Telefono_1
   Mensajes personalizados con Razon Social y Cartera

═══════════════════════════════════════════════════════════════════
  REQUISITOS DEL SISTEMA
═══════════════════════════════════════════════════════════════════

✓ Windows 10 o 11
✓ Python 3.11 o superior
✓ Celular Android con depuración USB
✓ Cable USB
✓ WhatsApp instalado en el celular

═══════════════════════════════════════════════════════════════════
  SOLUCIÓN DE PROBLEMAS
═══════════════════════════════════════════════════════════════════

PROBLEMA: No detecta dispositivos
SOLUCIÓN: 
   - Verifica que la depuración USB esté habilitada
   - Desconecta y vuelve a conectar el cable
   - Acepta el permiso en el celular

PROBLEMA: Error al leer CSV
SOLUCIÓN:
   - Verifica que el archivo tenga columnas de teléfono
   - Asegúrate de que use delimitador ; o ,
   - Prueba guardarlo como Excel (.xlsx)

PROBLEMA: No genera URLs
SOLUCIÓN:
   - Verifica que hayas seleccionado al menos un teléfono
   - Verifica que hayas seleccionado columnas para el mensaje
   - Verifica que hayas escrito una plantilla de mensaje

═══════════════════════════════════════════════════════════════════
  NOTAS IMPORTANTES
═══════════════════════════════════════════════════════════════════

• Los números se procesan automáticamente con prefijo 549
• Si una fila no tiene teléfono, se omite automáticamente
• Los desplegables inician cerrados para mejor organización
• Los botones solo muestran columnas seleccionadas
• El Excel procesado tiene SOLO URLs, sin datos personales
• Puedes pausar y reanudar el envío en cualquier momento

═══════════════════════════════════════════════════════════════════
  SOPORTE
═══════════════════════════════════════════════════════════════════

Para problemas o consultas, contacta a Berna.

═══════════════════════════════════════════════════════════════════

HERMES V1 - 2025 - Berna

