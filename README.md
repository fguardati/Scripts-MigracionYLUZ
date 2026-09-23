# Scripts-MigracionYLUZ

**Migración de flujos de Power Automate entre entornos.**

Guía de uso de `Script-CambiarSitiosListasShp.ps1` y consideraciones para que un flujo exportado como paquete `.zip` (formato heredado) se importe y funcione correctamente en el entorno o tenant de destino.

---

## 1. Qué hace el script

Trabaja sobre el archivo `definition.json` del paquete y genera un `new_definition.json` listo para volver a empaquetar:

| Qué cambia | Dónde | Con qué valor |
|---|---|---|
| GUID de la lista o biblioteca | `inputs.parameters.table` de cada operación de SharePoint | "GUID destino" del mapeo |
| URL del sitio | `inputs.parameters.dataset` de esa misma operación | "sitio destino" del mapeo |
| Conexiones "Invoker" *(opcional)* | `connectionReferences.*.source` y `inputs.authentication` | `Embedded` y `@parameters('$authentication')` |

Recorre el trigger y todas las acciones, incluidas las anidadas en Ámbitos, Condiciones (ramas Sí y No), Switch (casos y predeterminado), Aplicar a cada uno y Hasta.

**Nunca modifica** el `definition.json` ni el `list_mapping.csv` originales.

**No hace:**
- Reemplazar referencias de otros conectores (Teams, Outlook, Planner, Forms, Excel, etc.).
- Cambiar correos, URLs escritas dentro de expresiones o textos, nombres de columnas ni vistas.
- Descomprimir o volver a comprimir el `.zip`.

Ver [sección 7](#7-consideraciones-para-un-pasaje-correcto).

---

## 2. Requisitos

- Windows PowerShell 5.1 o PowerShell 7. Se recomienda **7.5 o superior**, o 5.1: las versiones 7.0 a 7.4 pueden cambiar el formato de las fechas del JSON (el script lo avisa).
- El script debe conservar la codificación **UTF-8 con BOM**. Si se edita y se guarda sin BOM, Windows PowerShell 5.1 muestra mal los acentos.
- Si PowerShell bloquea la ejecución por la directiva de ejecución, se puede habilitar solo para esa ejecución:

```bash
powershell -ExecutionPolicy Bypass -File .\Script-CambiarSitiosListasShp.ps1 -RutaDefinicion .\definition.json
```

---

## 3. Archivo de mapeo `list_mapping.csv`

- Codificación UTF-8, con fila de encabezado y **exactamente 4 campos** por línea, en este orden:

```
nombre_lista;GUID origen;GUID destino;sitio destino
Análisis químicos;c81ff860-5f23-41b4-a932-fe541508952b;0f8e2c1a-1b2c-4d5e-8f90-a1b2c3d4e5f6;https://destino.sharepoint.com/sites/quimicos
Usuarios;{4CEB6781-6D2C-4406-9273-7A2E3CD1A247};9a8b7c6d-5e4f-4a3b-2c1d-0e9f8a7b6c5d;https://destino.sharepoint.com/sites/quimicos
```

- Los 4 campos son obligatorios. Las líneas vacías se ignoran.
- Los GUIDs pueden ir en mayúsculas o minúsculas, con o sin llaves `{}`. El script los normaliza.
- `sitio destino` debe empezar con `https://`.
- No puede haber dos filas con el mismo GUID origen.
- El separador por defecto es `;`. Se cambia con `-Separador`.
- **No uses el separador dentro de los valores** (por ejemplo, un `;` en el nombre de la lista): el archivo se lee como texto plano, sin comillas.

### Cómo obtener el GUID de una lista

- En SharePoint: **Configuración de la lista**. El GUID aparece en la URL como `List=%7B...%7D` (`%7B` y `%7D` son las llaves).
- O directamente del `definition.json` de origen: si un GUID falta en el mapeo, el script lo informa (ver paso 2 del procedimiento).

---

## 4. Procedimiento paso a paso

### 4.1. Exportar el flujo de origen

**Mis flujos → … → Exportar → Paquete (.zip)**.

### 4.2. Descomprimir el paquete

El archivo a procesar está en:

```
Microsoft.Flow\flows\<GUID-del-flujo>\definition.json
```

### 4.3. Ejecutar el script

```bash
.\Script-CambiarSitiosListasShp.ps1 -RutaDefinicion ".\MiFlujo\Microsoft.Flow\flows\<GUID>\definition.json" -RutaMapeo .\list_mapping.csv
```

El script ejecuta estos pasos en orden y se detiene ante el primer problema:

| Paso | Qué hace | Si hay un problema |
|---|---|---|
| **0. Validación del CSV** | Cuenta los campos de cada línea y valida GUIDs, URLs y duplicados | Muestra la línea y el error. Cancela sin generar archivos |
| **1. Relevamiento** | Lista cada GUID encontrado, en qué acción está y si está mapeado. Detecta conexiones Invoker | — |
| **2. GUIDs faltantes** | Genera `list_mapping_pendiente.csv`: una copia del mapeo con una línea `;<guid>;;` por cada GUID faltante | Cancela sin generar `new_definition.json` |
| **3. Generación** | Reemplaza `table` y `dataset` (y las conexiones Invoker, si se indicó) y guarda `new_definition.json` | — |

Si se generó el archivo de pendientes: completá `nombre_lista`, `GUID destino` y `sitio destino` en las líneas nuevas, reemplazá con él el `list_mapping.csv` y volvé a ejecutar.

### 4.4. Si el flujo tiene conexiones "Invoker"

El script cancela con el código de salida 3. Leé la [sección 6](#6-conexiones-invoker) y, si corresponde, volvé a ejecutar agregando `-ConvertirInvokerAEmbedded`.

### 4.5. Reemplazar la definición en el paquete

1. Borrá el `definition.json` original de la carpeta del flujo.
2. Renombrá `new_definition.json` a **`definition.json`**. El paquete solo reconoce ese nombre.

### 4.6. Volver a comprimir (punto crítico)

Entrá a la carpeta descomprimida, **seleccioná `manifest.json` y la carpeta `Microsoft.Flow`** y hacé clic derecho → *Comprimir*.

```
✅ Correcto                              ❌ Incorrecto
MiFlujo.zip                              MiFlujo.zip
├── manifest.json                        └── MiFlujo/
└── Microsoft.Flow/                          ├── manifest.json
    └── flows/...                            └── Microsoft.Flow/...
```

- Si `manifest.json` no queda en la raíz del `.zip`, la importación falla.
- Usá el Explorador de Windows o 7-Zip. Evitá `Compress-Archive` de Windows PowerShell 5.1, que puede guardar las rutas con `\`.

### 4.7. Importar en destino

1. Entrá a **make.powerautomate.com** y verificá arriba a la derecha que estés en el **entorno de destino**.
2. Andá a **Mis flujos → Importar → Importar paquete (heredado)**.
3. En **Configuración de importación**, elegí *Crear como nuevo* o *Actualizar*, según corresponda.
4. Para cada conexión, **seleccioná o creá una conexión del tenant destino**.
5. Hacé clic en **Importar**.

### 4.8. Revisión posterior

Completá el [checklist de la sección 8](#8-checklist-posterior-a-la-importación).

---

## 5. Parámetros y códigos de salida

| Parámetro | Obligatorio | Por defecto | Descripción |
|---|---|---|---|
| `-RutaDefinicion` | Sí | — | Ruta del `definition.json` de entrada |
| `-RutaMapeo` | No | `list_mapping.csv` en la carpeta del `definition.json`; si no está, en la del script | Archivo de mapeo |
| `-Separador` | No | `;` | Separador de campos del CSV |
| `-CarpetaSalida` | No | Carpeta del `definition.json` | Dónde se generan `new_definition.json` y el CSV de pendientes. Se crea si no existe |
| `-ConvertirInvokerAEmbedded` | No | Desactivado | Convierte las conexiones Invoker a Embedded ([sección 6](#6-conexiones-invoker)) |

| Código | Significado |
|---|---|
| `0` | OK. Se generó `new_definition.json` |
| `1` | Error: falta un archivo, JSON inválido o CSV mal formado |
| `2` | Hay GUIDs sin mapear. Se generó `list_mapping_pendiente.csv` |
| `3` | Hay conexiones Invoker y no se indicó `-ConvertirInvokerAEmbedded` |

### Advertencias que puede mostrar el script

| Advertencia | Qué hacer |
|---|---|
| `'table' no es un GUID` | La lista se referencia por nombre o por una expresión. Revisarla a mano en el flujo importado |
| `'dataset' es una expresión` | El sitio viene de una variable o un parámetro. Se reemplaza el `table` pero el sitio queda como está. Verificar que la expresión resuelva al sitio destino |
| `usa una vista (view = ...)` | El GUID de la vista es distinto en destino. Volver a elegir la vista en la acción |
| `operación(es) de SharePoint sin parámetro 'table'` | Por ejemplo "Enviar solicitud HTTP a SharePoint". Revisar la URI a mano |

---

## 6. Conexiones "Invoker"

### Qué son

Son las conexiones de un flujo **instantáneo (de botón)** configuradas en *Usuarios que solo pueden ejecutar* como **"Proporcionada por el usuario que solo tiene permisos de ejecución"**. Cada acción toma la credencial de quien aprieta el botón.

### Por qué hay que convertirlas

Un paquete con conexiones Invoker **falla al importar** con este error:

```
{"error":{"code":"MissingAuthorizationHeaderAndClientCertificate","message":"A client certificate or authorization header was not provided."}}
```

Con `-ConvertirInvokerAEmbedded`, las acciones pasan a usar la conexión del propio flujo y el paquete se importa sin problemas.

### Qué impacto tiene si no se restaura después de importar

Mientras la conexión quede en *Embedded*, **todas las acciones se ejecutan con la cuenta de la conexión elegida al importar**, no con la de quien ejecuta el flujo:

| Aspecto | Efecto |
|---|---|
| **Permisos** ⚠️ | Cualquier usuario que pueda ejecutar el flujo lee y escribe con los permisos de esa cuenta, aunque él no tenga acceso a la lista. **Es una escalada de privilegios** |
| **Autoría y auditoría** | "Creado por", "Modificado por" y los registros de auditoría de SharePoint muestran la cuenta de la conexión |
| **Acciones "personales"** | Enviar correo, "mi perfil", "mi OneDrive", etc. se ejecutan como esa cuenta. Por ejemplo, los correos salen de su buzón |
| **Punto único de falla** | Si esa cuenta se da de baja, pierde la licencia o la conexión expira, el flujo deja de funcionar para todos |
| **Lo que no cambia** | Los datos del usuario que ejecuta, que vienen del trigger (por ejemplo, su correo), siguen disponibles |

### Cuándo restaurar el modo Invoker

| Si el flujo… | ¿Restaurar? |
|---|---|
| Depende de que cada usuario opere con sus propios permisos o de la trazabilidad por usuario | **Sí** |
| Lo ejecutan usuarios que igual tienen acceso a esas listas y no importa quién figure como autor | No es necesario |

### Cómo restaurarlo

1. Abrí el flujo importado → **Usuarios que solo pueden ejecutar** → **Editar**.
2. Para cada conexión que indicó el script en el bloque **"ACCIÓN REQUERIDA DESPUÉS DE IMPORTAR"**, elegí **"Proporcionada por el usuario que solo tiene permisos de ejecución"**.
3. Agregá los usuarios o grupos del tenant destino que deben poder ejecutarlo.
4. Guardá.

> Guardá la salida de consola del script o anotá qué flujos se convirtieron, para no perder el seguimiento de cuáles hay que restaurar.

---

## 7. Consideraciones para un pasaje correcto

Entre tenants, **nada de la identidad viaja**: ni usuarios, ni conexiones, ni permisos. Revisá estos puntos antes y después de migrar cada flujo.

### 7.1. Listas de SharePoint en destino

- **Nombres internos de las columnas.** El flujo usa nombres internos (`item/Legajo`, `item/Solo_visualizacion`, …), no los nombres visibles. Si las listas se crearon a mano en destino, los nombres internos pueden ser distintos aunque el nombre visible coincida: con acentos o espacios aparecen `_x00f3_`, `_x0020_` o `field_1`. **Creá las listas desde una plantilla o con PnP** para conservarlos.
- **Columnas de elección.** Los valores tienen que coincidir exactamente, incluidos los acentos y las mayúsculas.
- **Columnas de búsqueda (lookup).** Guardan IDs de elementos de otra lista, que pueden no coincidir en destino.
- **Columnas de tipo Persona.** Hacen referencia a usuarios del tenant de origen.
- **Vistas.** Tienen un GUID propio. Volvé a elegir la vista en cada acción que la use.

### 7.2. Referencias que el script no cambia

Buscá estos valores en el `new_definition.json`, o revisá el flujo importado:

- **Correos y UPN** (destinatarios, aprobadores, buzones compartidos). Si cambia el dominio, hay que actualizarlos.
- **Dominio del tenant de origen** (por ejemplo `ypf.sharepoint.com`) dentro de expresiones, `concat()`, links o la URI de "Enviar solicitud HTTP a SharePoint".
- **IDs de otros conectores:** Teams (`groupId`, `channelId`), Planner (plan, bucket), Forms (`formId`), OneDrive y Excel (unidad, archivo) y grupos de Entra ID.

### 7.3. Lo que el paquete no lleva

- **Conexiones.** Se crean o se eligen al importar, con cuentas del tenant destino.
- **Propietarios, copropietarios, usuarios de solo ejecución, historial de ejecuciones y alertas.**
- **Flujos secundarios (child flows) y variables de entorno.** Requieren soluciones, no paquetes heredados.
- **Accesos a los flujos de botón.** El menú "Automatizar" de una lista, los botones de Power Apps y la app móvil se tienen que volver a vincular.

### 7.4. Políticas del tenant destino

- **DLP:** las directivas de prevención de pérdida de datos pueden bloquear combinaciones de conectores (por ejemplo, HTTP junto con SharePoint).
- **Licencias:** los conectores premium requieren licencia para el propietario o para el flujo.

---

## 8. Checklist posterior a la importación

- [ ] El flujo está **activado**.
- [ ] Cada conexión apunta a una cuenta del **tenant destino**.
- [ ] Los sitios y listas de cada acción de SharePoint se ven bien en el diseñador, sin errores de "lista no encontrada".
- [ ] Las columnas de las acciones "Crear elemento" y "Actualizar elemento" aparecen con sus campos, sin campos vacíos ni faltantes.
- [ ] Se volvieron a elegir las vistas indicadas por el script.
- [ ] Se actualizaron los correos, los dominios y los IDs de otros conectores ([sección 7.2](#72-referencias-que-el-script-no-cambia)).
- [ ] Si hubo conversión Invoker → Embedded: se restauró la configuración o se decidió dejarla así a conciencia ([sección 6](#6-conexiones-invoker)).
- [ ] Se agregaron propietarios y usuarios de solo ejecución del tenant destino.
- [ ] Se hizo una **ejecución de prueba** y se verificó el resultado en SharePoint.

---

## 9. Problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| `MissingAuthorizationHeaderAndClientCertificate` al importar | Conexiones Invoker en el paquete | Ejecutar con `-ConvertirInvokerAEmbedded` |
| El mismo error, sin conexiones Invoker | Sesión del portal vencida | Recargar la página, volver a iniciar sesión o probar en una ventana InPrivate |
| La importación falla o no reconoce el paquete | `manifest.json` no está en la raíz del `.zip` | Volver a comprimir seleccionando el contenido ([4.6](#46-volver-a-comprimir-punto-crítico)) |
| El paquete importa pero conserva los valores viejos | Quedó `new_definition.json` sin renombrar | Renombrarlo a `definition.json` ([4.5](#45-reemplazar-la-definición-en-el-paquete)) |
| `Línea N: se encontraron X campo(s), se esperaban 4` | Falta un campo, sobra un separador o el separador es otro | Corregir la línea o usar `-Separador` |
| Acentos mal mostrados en consola (5.1) | El script se guardó sin BOM | Volver a guardarlo como UTF-8 con BOM |
| En el JSON aparecen `'`, `<` | Serialización de Windows PowerShell 5.1 | Es JSON válido y equivalente. No requiere acción |
| Fechas con otro formato en el JSON | PowerShell 7.0 a 7.4 | Usar PowerShell 7.5+ o 5.1 |
