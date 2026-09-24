<#
.SYNOPSIS
    Prepara la migración de un flujo de Power Automate entre entornos, reemplazando
    las referencias a listas y sitios de SharePoint en definition.json.

.DESCRIPTION
    Trabaja sobre el archivo definition.json de un paquete .zip de flujo (formato heredado).
    Usa un archivo de mapeo (list_mapping.csv) mantenido manualmente para reemplazar, en cada
    operación del conector de SharePoint (trigger o acción, incluidas las anidadas en
    Condiciones, Aplicar a cada uno, Ámbitos, Switch, Hasta, etc.):
      - inputs.parameters.table   (GUID de la lista)  -> GUID destino
      - inputs.parameters.dataset (URL del sitio)     -> sitio destino

    Soporta los dos formatos de acción de conector, que pueden mezclarse en un mismo flujo
    (el formato se decide por operación según su "type"):
      - actual  (OpenApiConnection*): sitio y lista en inputs.parameters.dataset / table.
      - clásico (ApiConnection*, heredado de Logic Apps): sitio y lista embebidos en el
        string inputs.path, p. ej.
        /datasets/@{encodeURIComponent(encodeURIComponent('<sitio>'))}/tables/@{encodeURIComponent(encodeURIComponent('<GUID>'))}/items

    Pasos:
      0. Valida la estructura del CSV (exactamente 4 campos por registro).
      1. Releva los GUIDs de "table" y los compara contra la columna "GUID origen".
      2. Si hay GUIDs sin mapear, genera <mapeo>_pendiente.csv y cancela.
      3. Genera new_definition.json con los reemplazos (UTF-8 sin BOM).

    Conexiones "Invoker" (flujos instantáneos con la conexión configurada como
    "Proporcionada por el usuario que solo tiene permisos de ejecución"):
    un paquete con este tipo de conexión falla al importar con el error
    MissingAuthorizationHeaderAndClientCertificate. El script las detecta y:
      - sin -ConvertirInvokerAEmbedded: informa y cancela sin generar new_definition.json.
      - con -ConvertirInvokerAEmbedded: las convierte a "Embedded" (la conexión del flujo)
        y avisa que hay que restaurar la configuración después de importar.

    Los archivos definition.json y list_mapping.csv originales nunca se modifican.

    Formato de list_mapping.csv (UTF-8, con encabezado, 4 campos obligatorios):
      nombre_lista;GUID origen;GUID destino;sitio destino

.PARAMETER RutaDefinicion
    Ruta del archivo definition.json de entrada.

.PARAMETER RutaMapeo
    Ruta del archivo de mapeo. Por defecto: list_mapping.csv en la carpeta del
    definition.json (si no existe ahí, se busca en la carpeta del script).

.PARAMETER Separador
    Separador de campos del CSV. Por defecto ";".

.PARAMETER CarpetaSalida
    Carpeta donde se generan new_definition.json y el CSV de pendientes.
    Por defecto: la carpeta del definition.json. Se crea si no existe.

.PARAMETER ConvertirInvokerAEmbedded
    Convierte las conexiones "Invoker" a "Embedded" para que el paquete se pueda importar.
    Mientras no se restaure la configuración en el flujo importado, TODAS las acciones se
    ejecutan con la conexión elegida al importar (sus permisos y su identidad), no con la
    de cada usuario que ejecuta el flujo.

.EXAMPLE
    .\Script-CambiarSitiosListasShp.ps1 -RutaDefinicion "C:\Export\Microsoft.Flow\flows\0a1b2c3d-...\definition.json"

    Usa list_mapping.csv con separador ";" y deja new_definition.json junto al original.

.EXAMPLE
    .\Script-CambiarSitiosListasShp.ps1 -RutaDefinicion .\definition.json -RutaMapeo .\mapeos\prod.csv -Separador "," -CarpetaSalida .\salida

.EXAMPLE
    .\Script-CambiarSitiosListasShp.ps1 -RutaDefinicion .\definition.json -ConvertirInvokerAEmbedded

    Además de los reemplazos de SharePoint, convierte las conexiones Invoker a Embedded.

.NOTES
    Compatible con Windows PowerShell 5.1 y PowerShell 7.
    Códigos de salida: 0 = OK, 1 = error, 2 = hay GUIDs pendientes de mapear,
                       3 = hay conexiones Invoker y no se indicó -ConvertirInvokerAEmbedded.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$RutaDefinicion,

    [string]$RutaMapeo,

    [string]$Separador = ';',

    [string]$CarpetaSalida,

    [switch]$ConvertirInvokerAEmbedded
)

$ErrorActionPreference = 'Stop'

# Expresión regular para el formato clásico: captura el sitio y la lista dentro de inputs.path.
# (?:encodeURIComponent\()+ acepta una o más capas de encodeURIComponent. Lo que queda fuera
# del match (/items, /items/@{...}/attachments, /onnewitems, etc.) no se modifica.
$PatronPathClasico = @'
(?<pre>/datasets/@\{(?:encodeURIComponent\()+')(?<site>[^']*)(?<mid>'\)+\}/tables/@\{(?:encodeURIComponent\()+')(?<table>[^']*)(?<post>'\)+\})
'@.Trim()

#region Utilidades de consola y acceso a datos

function Write-Titulo([string]$Texto) {
    Write-Host ''
    Write-Host "=== $Texto ===" -ForegroundColor Cyan
}

function Write-Ok([string]$Texto)    { Write-Host "[OK] $Texto" -ForegroundColor Green }
function Write-Aviso([string]$Texto) { Write-Host "[ADVERTENCIA] $Texto" -ForegroundColor Yellow }
function Write-Fallo([string]$Texto) { Write-Host "[ERROR] $Texto" -ForegroundColor Red }

# Devuelve el valor de una propiedad de un objeto JSON, o $null si no existe.
# Se usa "return ," para que los arrays no se desenrollen al salir de la función.
function Get-Propiedad($Objeto, [string]$Nombre) {
    if ($Objeto -isnot [System.Management.Automation.PSCustomObject]) { return $null }
    $prop = $Objeto.PSObject.Properties[$Nombre]
    if ($null -eq $prop) { return $null }
    return , $prop.Value
}

# Normaliza un GUID: quita espacios y llaves, valida el formato y lo devuelve en
# minúsculas (formato xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx). Si no es un GUID devuelve $null.
function ConvertTo-GuidNormalizado($Valor) {
    if ($Valor -isnot [string] -or [string]::IsNullOrWhiteSpace($Valor)) { return $null }
    $limpio = $Valor.Trim().TrimStart('{').TrimEnd('}').Trim()
    $guid = [guid]::Empty
    if ([guid]::TryParseExact($limpio, 'D', [ref]$guid)) { return $guid.ToString('D') }
    return $null
}

# Convierte una ruta (relativa o absoluta) en ruta completa del sistema de archivos.
# Necesario porque los métodos de [System.IO.File] no usan la ubicación actual de PowerShell.
function Resolve-RutaCompleta([string]$Ruta) {
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Ruta)
}

#endregion

#region Paso 0: Validación del CSV de mapeo

# Valida la estructura del CSV leyéndolo como texto plano: cada registro (incluido el
# encabezado) debe tener exactamente 4 campos. Las líneas vacías se ignoran.
# Si hay errores, los muestra todos y lanza una excepción para cancelar la ejecución.
function Test-EstructuraMapeo([string]$Ruta, [string]$Separador) {
    $lineas  = [System.IO.File]::ReadAllLines($Ruta, [System.Text.Encoding]::UTF8)
    $errores = New-Object System.Collections.Generic.List[string]
    $registros = 0

    for ($i = 0; $i -lt $lineas.Length; $i++) {
        if ([string]::IsNullOrWhiteSpace($lineas[$i])) { continue }
        $registros++
        $campos = $lineas[$i].Split([string[]]@($Separador), [System.StringSplitOptions]::None)
        if ($campos.Length -ne 4) {
            $errores.Add(("Línea {0}: se encontraron {1} campo(s), se esperaban 4." -f ($i + 1), $campos.Length))
        }
    }

    if ($registros -eq 0) {
        throw "El archivo de mapeo está vacío (no tiene ni siquiera el encabezado): $Ruta"
    }
    if ($errores.Count -gt 0) {
        foreach ($e in $errores) { Write-Fallo $e }
        throw ("El archivo de mapeo tiene {0} registro(s) con una cantidad de campos incorrecta (separador '{1}'). No se generó ningún archivo." -f $errores.Count, $Separador)
    }

    Write-Ok ("Estructura del CSV válida: {0} registro(s) con 4 campos (incluido el encabezado)." -f $registros)
}

# Lee el contenido del mapeo (ya validado estructuralmente) y verifica que los campos
# obligatorios estén completos y tengan formato correcto. Devuelve un hashtable
# indexado por GUID origen normalizado.
function Read-Mapeo([string]$Ruta, [string]$Separador) {
    $lineas  = [System.IO.File]::ReadAllLines($Ruta, [System.Text.Encoding]::UTF8)
    $errores = New-Object System.Collections.Generic.List[string]
    $mapeo   = @{}
    $esEncabezado = $true

    for ($i = 0; $i -lt $lineas.Length; $i++) {
        if ([string]::IsNullOrWhiteSpace($lineas[$i])) { continue }
        if ($esEncabezado) { $esEncabezado = $false; continue }

        $nro    = $i + 1
        $campos = $lineas[$i].Split([string[]]@($Separador), [System.StringSplitOptions]::None) | ForEach-Object { $_.Trim() }
        $nombreLista = $campos[0]
        $origen  = ConvertTo-GuidNormalizado $campos[1]
        $destino = ConvertTo-GuidNormalizado $campos[2]
        $sitio   = $campos[3]

        $nombres = @('nombre_lista', 'GUID origen', 'GUID destino', 'sitio destino')
        $vacios  = @(for ($c = 0; $c -lt 4; $c++) { if ($campos[$c] -eq '') { $nombres[$c] } })
        if ($vacios.Count -gt 0) {
            $errores.Add(("Línea {0}: campo(s) obligatorio(s) vacío(s): {1}." -f $nro, ($vacios -join ', ')))
            continue
        }
        if (-not $origen)  { $errores.Add("Línea ${nro}: 'GUID origen' no es un GUID válido ('$($campos[1])').") }
        if (-not $destino) { $errores.Add("Línea ${nro}: 'GUID destino' no es un GUID válido ('$($campos[2])').") }
        if ($sitio -notmatch '^https?://') { $errores.Add("Línea ${nro}: 'sitio destino' no es una URL válida ('$sitio').") }
        if ($origen -and $mapeo.ContainsKey($origen)) {
            $errores.Add(("Línea {0}: el GUID origen {1} ya está mapeado en la línea {2}." -f $nro, $origen, $mapeo[$origen].Linea))
        }
        if (-not $origen -or -not $destino -or $sitio -notmatch '^https?://' -or $mapeo.ContainsKey($origen)) { continue }

        $mapeo[$origen] = [pscustomobject]@{
            Linea        = $nro
            NombreLista  = $nombreLista
            GuidDestino  = $destino
            SitioDestino = $sitio
        }
    }

    if ($errores.Count -gt 0) {
        foreach ($e in $errores) { Write-Fallo $e }
        throw ("El archivo de mapeo tiene {0} error(es) de contenido. No se generó ningún archivo." -f $errores.Count)
    }
    if ($mapeo.Count -eq 0) { Write-Aviso 'El archivo de mapeo no tiene filas de datos (solo encabezado).' }
    else { Write-Ok ("Mapeo cargado: {0} lista(s)." -f $mapeo.Count) }

    return $mapeo
}

#endregion

#region Paso 1: Relevamiento de operaciones de SharePoint

# Lee y parsea definition.json. En PowerShell 7.5+ se usa -DateKind String para que las
# fechas se conserven como texto y no se reescriban al volver a serializar.
function Read-Definicion([string]$Ruta) {
    $texto = [System.IO.File]::ReadAllText($Ruta, [System.Text.Encoding]::UTF8)
    $parametros = @{ InputObject = $texto }
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) {
        $parametros['DateKind'] = 'String'
    }
    elseif ($PSVersionTable.PSVersion.Major -ge 6) {
        Write-Aviso 'Esta versión de PowerShell 7 (< 7.5) convierte las fechas del JSON a DateTime; su formato puede cambiar en new_definition.json. Se recomienda PowerShell 7.5+ o Windows PowerShell 5.1.'
    }
    try {
        return ConvertFrom-Json @parametros
    }
    catch {
        throw "El archivo no contiene un JSON válido: $Ruta`n        Detalle: $($_.Exception.Message)"
    }
}

# Ubica el nodo de la definición del flujo (el que contiene triggers y actions).
function Get-NodoDefinicion($Raiz) {
    $def = Get-Propiedad (Get-Propiedad $Raiz 'properties') 'definition'
    if ($null -eq $def) { $def = Get-Propiedad $Raiz 'definition' }
    if ($null -eq $def -and ((Get-Propiedad $Raiz 'triggers') -or (Get-Propiedad $Raiz 'actions'))) { $def = $Raiz }
    if ($null -eq $def) { throw 'No se encontró la definición del flujo (properties.definition) en el JSON.' }
    return $def
}

# Devuelve los nombres de las referencias de conexión que apuntan al conector de SharePoint,
# para reconocer acciones cuyo connectionName no contenga "sharepointonline".
function Get-ReferenciasSharePoint($Raiz) {
    $claves = @()
    $refs = Get-Propiedad (Get-Propiedad $Raiz 'properties') 'connectionReferences'
    if ($refs -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $refs.PSObject.Properties) {
            $api = Get-Propiedad $p.Value 'api'
            $texto = '{0} {1} {2}' -f (Get-Propiedad $p.Value 'id'), (Get-Propiedad $api 'name'), (Get-Propiedad $p.Value 'apiName')
            if ($texto -match 'sharepointonline') { $claves += $p.Name }
        }
    }
    return , $claves
}

# Indica si una acción o trigger usa el conector de SharePoint (según inputs.host).
function Test-EsOperacionSharePoint($Accion, [string[]]$ReferenciasSP) {
    $hostAccion = Get-Propiedad (Get-Propiedad $Accion 'inputs') 'host'
    if ($null -eq $hostAccion) { return $false }
    $conexion = Get-Propiedad $hostAccion 'connectionName'
    $texto = '{0} {1} {2}' -f (Get-Propiedad $hostAccion 'apiId'), $conexion, (Get-Propiedad (Get-Propiedad $hostAccion 'connection') 'name')
    if ($texto -match 'sharepointonline') { return $true }
    return ($conexion -and $ReferenciasSP -contains $conexion)
}

# Clasifica una acción según su "type":
#   'clásico' -> ApiConnection, ApiConnectionWebhook, ApiConnectionNotification
#   'actual'  -> OpenApiConnection, OpenApiConnectionWebhook, OpenApiConnectionNotification
#   $null     -> no es una operación de conector (sus acciones internas se recorren igual)
function Get-FormatoOperacion($Accion) {
    $tipo = Get-Propiedad $Accion 'type'
    if (@('ApiConnection', 'ApiConnectionWebhook', 'ApiConnectionNotification') -contains $tipo) { return 'clásico' }
    if (@('OpenApiConnection', 'OpenApiConnectionWebhook', 'OpenApiConnectionNotification') -contains $tipo) { return 'actual' }
    return $null
}

# Indica si una operación en formato clásico usa el conector de SharePoint:
#   1. Extrae CLAVE de host.connection.name = @parameters('$connections')['CLAVE']['connectionId'].
#   2. Es SharePoint si properties.connectionReferences.CLAVE tiene apiName = "sharepointonline".
#   3. Respaldo, si la referencia no se puede resolver: host.api.runtimeUrl termina en "/sharepointonline".
function Test-EsSharePointClasico($Accion, $ReferenciasConexion) {
    $hostAccion = Get-Propiedad (Get-Propiedad $Accion 'inputs') 'host'
    $nombreConexion = Get-Propiedad (Get-Propiedad $hostAccion 'connection') 'name'
    if ($nombreConexion -is [string] -and
        $nombreConexion -match '^@parameters\(''\$connections''\)\[''(?<clave>[^'']+)''\]\[''connectionId''\]$') {
        $apiName = Get-Propiedad (Get-Propiedad $ReferenciasConexion $Matches['clave']) 'apiName'
        if ($apiName) { return ($apiName -eq 'sharepointonline') }
    }
    $runtimeUrl = Get-Propiedad (Get-Propiedad $hostAccion 'api') 'runtimeUrl'
    return ($runtimeUrl -is [string] -and $runtimeUrl -like '*/sharepointonline')
}

# Construye el registro de una operación de SharePoint (o $null si la acción no lo es),
# extrayendo sitio y lista según el formato. "Estado" indica si se puede procesar:
#   OK          -> GUID de lista válido; se compara contra el mapeo y se reemplaza
#   NoGuid      -> "table" literal que no es un GUID (p. ej. nombre de la lista)
#   SinTabla    -> la operación no referencia una lista (sin "table" / path sin /datasets/)
#   Vacio       -> sitio o lista vacíos
#   Dinamico    -> sitio o lista calculados con una expresión
#   SinTables   -> (clásico) path con /datasets/ pero sin /tables/
#   Incoherente -> la estructura no coincide con el tipo (sin inputs.path / inputs.parameters)
# Solo las operaciones con Estado = OK llevan Guid, así que las demás no cuentan como
# faltantes ni se modifican.
function Get-OperacionSharePoint($Item, [string[]]$ReferenciasSP, $ReferenciasConexion) {
    $accion  = $Item.Accion
    $formato = Get-FormatoOperacion $accion
    if (-not $formato) { return $null }

    $esSharePoint = if ($formato -eq 'clásico') { Test-EsSharePointClasico $accion $ReferenciasConexion }
                    else { Test-EsOperacionSharePoint $accion $ReferenciasSP }
    if (-not $esSharePoint) { return $null }

    $op = [pscustomobject]@{
        Nombre     = $Item.Nombre
        Ruta       = $Item.Ruta
        Tipo       = $Item.Tipo
        Formato    = $formato
        Accion     = $accion       # referencia al objeto en memoria, se modifica en el paso 3
        Parametros = $null
        TieneTabla = $false
        Tabla      = $null
        Sitio      = $null
        Guid       = $null
        Estado     = 'OK'
    }
    $inputs = Get-Propiedad $accion 'inputs'

    if ($formato -eq 'clásico') {
        $path = Get-Propiedad $inputs 'path'
        if ($path -isnot [string]) { $op.Estado = 'Incoherente'; return $op }

        if ($path -notlike '*/datasets/*') { $op.Estado = 'SinTabla'; return $op }
        if ($path -notlike '*/tables/*')   { $op.Estado = 'SinTables'; $op.Sitio = $path; return $op }

        $m = [regex]::Match($path, $PatronPathClasico)
        if (-not $m.Success) { $op.Estado = 'Dinamico'; $op.Sitio = $path; return $op }

        $op.TieneTabla = $true
        $op.Sitio = $m.Groups['site'].Value
        $op.Tabla = $m.Groups['table'].Value
    }
    else {
        $parametros = Get-Propiedad $inputs 'parameters'
        if ($parametros -isnot [System.Management.Automation.PSCustomObject]) { $op.Estado = 'Incoherente'; return $op }

        $op.Parametros = $parametros
        $op.TieneTabla = ($null -ne $parametros.PSObject.Properties['table'])
        $op.Tabla      = Get-Propiedad $parametros 'table'
        $op.Sitio      = Get-Propiedad $parametros 'dataset'
        if (-not $op.TieneTabla) { $op.Estado = 'SinTabla'; return $op }
    }

    # Casos especiales comunes a ambos formatos
    $valores = @($op.Sitio, $op.Tabla) | Where-Object { $_ -is [string] }
    if (@($valores | Where-Object { $_.Trim() -eq '' }).Count -gt 0) { $op.Estado = 'Vacio'; return $op }
    if (@($valores | Where-Object { $_.StartsWith('@') }).Count -gt 0) { $op.Estado = 'Dinamico'; return $op }

    $op.Guid = ConvertTo-GuidNormalizado $op.Tabla
    if (-not $op.Guid) { $op.Estado = 'NoGuid' }
    return $op
}

# Recorre recursivamente una colección de acciones (o triggers) y agrega a $Resultado cada
# una con su nombre, ubicación y tipo. Desciende en:
#   actions (Ámbito, Aplicar a cada uno, Hasta, rama "Sí" de Condición),
#   else.actions (rama "No"), cases.*.actions y default.actions (Switch).
function Add-Acciones($Coleccion, [string]$Ruta, [string]$Tipo, $Resultado) {
    if ($Coleccion -isnot [System.Management.Automation.PSCustomObject]) { return }

    foreach ($p in $Coleccion.PSObject.Properties) {
        $nombre = $p.Name
        $accion = $p.Value
        $rutaActual = if ($Ruta) { "$Ruta > $nombre" } else { $nombre }

        $Resultado.Add([pscustomobject]@{
            Nombre = $nombre
            Ruta   = $rutaActual
            Tipo   = $Tipo
            Accion = $accion   # referencia al objeto en memoria, se modifica en el paso 3
        })

        # Descenso en acciones contenedoras
        $tipoAccion = Get-Propiedad $accion 'type'
        $etiquetaSi = if ($tipoAccion -eq 'If') { "$rutaActual [Sí]" } else { $rutaActual }
        Add-Acciones (Get-Propiedad $accion 'actions') $etiquetaSi 'Acción' $Resultado
        Add-Acciones (Get-Propiedad (Get-Propiedad $accion 'else') 'actions') "$rutaActual [No]" 'Acción' $Resultado

        $casos = Get-Propiedad $accion 'cases'
        if ($casos -is [System.Management.Automation.PSCustomObject]) {
            foreach ($caso in $casos.PSObject.Properties) {
                Add-Acciones (Get-Propiedad $caso.Value 'actions') "$rutaActual [Caso: $($caso.Name)]" 'Acción' $Resultado
            }
        }
        Add-Acciones (Get-Propiedad (Get-Propiedad $accion 'default') 'actions') "$rutaActual [Predeterminado]" 'Acción' $Resultado
    }
}

# Devuelve todos los triggers y acciones del flujo (incluidas las anidadas).
function Get-TodasLasAcciones($Raiz) {
    $definicion = Get-NodoDefinicion $Raiz
    $acciones = New-Object System.Collections.Generic.List[object]
    Add-Acciones (Get-Propiedad $definicion 'triggers') '' 'Trigger' $acciones
    Add-Acciones (Get-Propiedad $definicion 'actions')  '' 'Acción'  $acciones
    return , $acciones
}

# Paso 1 completo: releva las operaciones de SharePoint, muestra el resumen y devuelve las
# operaciones junto con la lista de GUIDs (distintos) que no están en el mapeo.
function Get-RelevamientoGuids($Raiz, $Acciones, [hashtable]$Mapeo) {
    $refsSP      = Get-ReferenciasSharePoint $Raiz
    $refsConexion = Get-Propiedad (Get-Propiedad $Raiz 'properties') 'connectionReferences'
    $operaciones = New-Object System.Collections.Generic.List[object]
    foreach ($a in $Acciones) {
        $op = Get-OperacionSharePoint $a $refsSP $refsConexion
        if ($op) { $operaciones.Add($op) }
    }

    $faltantes = New-Object System.Collections.Generic.List[string]
    $conGuid   = @($operaciones | Where-Object { $_.Guid })

    if ($conGuid.Count -eq 0) {
        Write-Aviso 'No se encontraron operaciones de SharePoint con un GUID de lista en "table".'
    }
    foreach ($op in $conGuid) {
        $ubicacion = if ($op.Ruta -ne $op.Nombre) { "  (en: $($op.Ruta))" } else { '' }
        if ($Mapeo.ContainsKey($op.Guid)) {
            Write-Host ("  [MAPEADO]    {0}  {1} ({5}): {2}  -> lista '{3}'{4}" -f $op.Guid, $op.Tipo, $op.Nombre, $Mapeo[$op.Guid].NombreLista, $ubicacion, $op.Formato) -ForegroundColor Green
        }
        else {
            Write-Host ("  [SIN MAPEO]  {0}  {1} ({4}): {2}{3}" -f $op.Guid, $op.Tipo, $op.Nombre, $ubicacion, $op.Formato) -ForegroundColor Yellow
            if (-not $faltantes.Contains($op.Guid)) { $faltantes.Add($op.Guid) }
        }
    }

    # Operaciones de SharePoint que no se procesan: se advierten y no se modifican
    foreach ($op in @($operaciones | Where-Object { $_.Estado -notin 'OK', 'SinTabla' })) {
        $prefijo = "{0} '{1}' (formato {2})" -f $op.Tipo, $op.Ruta, $op.Formato
        switch ($op.Estado) {
            'NoGuid' {
                Write-Aviso ("{0}: 'table' no es un GUID ('{1}'); no se reemplazará. Revisar manualmente." -f $prefijo, $op.Tabla)
            }
            'Incoherente' {
                $esperado = if ($op.Formato -eq 'clásico') { 'inputs.path' } else { 'inputs.parameters' }
                Write-Aviso ("{0}: el tipo '{1}' debería tener {2} y no lo tiene; no se procesa." -f $prefijo, $op.Accion.type, $esperado)
            }
            'Vacio' {
                Write-Aviso ("{0}: el sitio o la lista están vacíos (sitio = '{1}', lista = '{2}'); no se procesa." -f $prefijo, $op.Sitio, $op.Tabla)
            }
            'Dinamico' {
                $valor = if ($op.Formato -eq 'clásico') { "path = $($op.Sitio)" } else { "dataset = '$($op.Sitio)', table = '$($op.Tabla)'" }
                Write-Aviso ("{0}: valor dinámico, revisar manualmente ({1})." -f $prefijo, $valor)
            }
            'SinTables' {
                Write-Aviso ("{0}: el path tiene /datasets/ pero no /tables/ (p. ej. HTTP u operación de archivos); revisar manualmente (path = {1})." -f $prefijo, $op.Sitio)
            }
        }
    }
    $sinTabla = @($operaciones | Where-Object { $_.Estado -eq 'SinTabla' }).Count
    if ($sinTabla -gt 0) {
        Write-Host ("  Info: {0} operación(es) de SharePoint sin parámetro 'table' (p. ej. HTTP a SharePoint); no se modifican." -f $sinTabla) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host ("  Operaciones con GUID: {0} | Mapeadas: {1} | GUIDs distintos sin mapeo: {2}" -f $conGuid.Count, ($conGuid.Count - @($conGuid | Where-Object { -not $Mapeo.ContainsKey($_.Guid) }).Count), $faltantes.Count)

    return [pscustomobject]@{
        Operaciones = $operaciones
        Faltantes   = $faltantes
    }
}

#endregion

#region Paso 1 (bis): Detección de conexiones "Invoker"

# Indica si una acción toma la conexión del usuario que ejecuta el flujo, es decir, si su
# autenticación sale de los encabezados del trigger (X-MS-APIM-Tokens / $ConnectionKey).
function Test-AutenticacionInvoker($Accion) {
    $auth  = Get-Propiedad (Get-Propiedad $Accion 'inputs') 'authentication'
    $valor = if ($auth -is [string]) { $auth } else { Get-Propiedad $auth 'value' }
    return ($valor -is [string] -and $valor -match 'X-MS-APIM-Tokens')
}

# Detecta la configuración "Proporcionada por el usuario que solo tiene permisos de ejecución":
#   - referencias de conexión con "source": "Invoker"
#   - acciones cuya autenticación depende del usuario que ejecuta
# Muestra el resultado en consola y lo devuelve para el paso 3.
function Find-ConexionesInvoker($Raiz, $Acciones, [bool]$Convertir) {
    $referencias = New-Object System.Collections.Generic.List[object]
    $refs = Get-Propiedad (Get-Propiedad $Raiz 'properties') 'connectionReferences'
    if ($refs -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $refs.PSObject.Properties) {
            if ((Get-Propiedad $p.Value 'source') -eq 'Invoker') {
                $referencias.Add([pscustomobject]@{ Nombre = $p.Name; Referencia = $p.Value })
            }
        }
    }
    $accionesInvoker = @($Acciones | Where-Object { Test-AutenticacionInvoker $_.Accion })
    $hay = ($referencias.Count + $accionesInvoker.Count) -gt 0

    if ($hay) {
        Write-Host ''
        Write-Aviso "Se detectaron conexiones configuradas como 'Proporcionada por el usuario que solo tiene permisos de ejecución' (Invoker):"
        foreach ($r in $referencias) { Write-Host "    - Referencia de conexión: $($r.Nombre)" -ForegroundColor Yellow }
        foreach ($a in $accionesInvoker) {
            $ubicacion = if ($a.Ruta -ne $a.Nombre) { "  (en: $($a.Ruta))" } else { '' }
            Write-Host ("    - {0}: {1}{2}" -f $a.Tipo, $a.Nombre, $ubicacion) -ForegroundColor Yellow
        }
        if ($Convertir) {
            Write-Host '    Se convertirán a Embedded en el paso 3 (-ConvertirInvokerAEmbedded).' -ForegroundColor Yellow
        }
        else {
            Write-Aviso 'Un paquete con conexiones Invoker falla al importar (MissingAuthorizationHeaderAndClientCertificate).'
        }
    }
    elseif ($Convertir) {
        Write-Host '  Info: no se encontraron conexiones Invoker; -ConvertirInvokerAEmbedded no tendrá efecto.' -ForegroundColor DarkGray
    }

    return [pscustomobject]@{
        Hay         = $hay
        Referencias = $referencias
        Acciones    = $accionesInvoker
    }
}

#endregion

#region Paso 2: Generación del CSV de pendientes

# Genera una copia idéntica del CSV de mapeo (misma codificación/BOM y fin de línea) con una
# línea adicional por cada GUID faltante: ";<guid>;;" (4 campos, solo GUID origen completo).
function New-MapeoPendiente([string]$RutaMapeo, [string]$Separador, [string[]]$Faltantes, [string]$CarpetaSalida) {
    $bytes    = [System.IO.File]::ReadAllBytes($RutaMapeo)
    $tieneBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $texto    = [System.IO.File]::ReadAllText($RutaMapeo, [System.Text.Encoding]::UTF8)
    $finLinea = if ($texto.Contains("`r`n")) { "`r`n" } else { "`n" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($texto)
    if ($texto.Length -gt 0 -and -not $texto.EndsWith("`n")) { [void]$sb.Append($finLinea) }
    foreach ($guid in $Faltantes) {
        [void]$sb.Append((@('', $guid, '', '') -join $Separador)).Append($finLinea)
    }

    $nombre = [System.IO.Path]::GetFileNameWithoutExtension($RutaMapeo) + '_pendiente.csv'
    $rutaPendiente = Join-Path $CarpetaSalida $nombre
    [System.IO.File]::WriteAllText($rutaPendiente, $sb.ToString(), (New-Object System.Text.UTF8Encoding($tieneBom)))
    return $rutaPendiente
}

#endregion

#region Paso 3: Reemplazo y generación de new_definition.json

# Aplica los reemplazos sobre el objeto en memoria (el archivo original no se toca) y
# devuelve la lista de cambios realizados.
function Update-ReferenciasSharePoint($Operaciones, [hashtable]$Mapeo) {
    $cambios = New-Object System.Collections.Generic.List[object]

    # Formato clásico: reconstruye cada match como pre + sitio destino + mid + GUID destino + post,
    # tomando sitio y GUID de la fila del mapeo cuyo GUID origen coincide con el grupo "table".
    $evaluador = [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $guid = ConvertTo-GuidNormalizado $m.Groups['table'].Value
        if (-not $guid -or -not $Mapeo.ContainsKey($guid)) { return $m.Value }
        $filaMapeo = $Mapeo[$guid]
        return $m.Groups['pre'].Value + $filaMapeo.SitioDestino + $m.Groups['mid'].Value + $filaMapeo.GuidDestino + $m.Groups['post'].Value
    }

    foreach ($op in @($Operaciones | Where-Object { $_.Guid -and $Mapeo.ContainsKey($_.Guid) })) {
        $fila   = $Mapeo[$op.Guid]

        if ($op.Formato -eq 'clásico') {
            $inputs = $op.Accion.inputs
            $inputs.path = [regex]::Replace($inputs.path, $PatronPathClasico, $evaluador)
            $cambios.Add([pscustomobject]@{
                Tipo            = $op.Tipo
                Nombre          = $op.Nombre
                Ruta            = $op.Ruta
                Lista           = $fila.NombreLista
                TablaAnterior   = $op.Tabla
                TablaNueva      = $fila.GuidDestino
                DatasetAnterior = $op.Sitio
                DatasetNuevo    = $fila.SitioDestino
            })
            continue
        }

        # Formato actual: asignación directa de inputs.parameters.dataset / table
        $params = $op.Parametros

        $tablaAnterior = $params.table
        $params.table  = $fila.GuidDestino

        $propDataset = $params.PSObject.Properties['dataset']
        $datasetAnterior = if ($propDataset) { $propDataset.Value } else { $null }
        if ($null -eq $propDataset) {
            $params | Add-Member -NotePropertyName 'dataset' -NotePropertyValue $fila.SitioDestino
            Write-Aviso ("{0} '{1}': no tenía 'dataset'; se agregó con el sitio destino." -f $op.Tipo, $op.Ruta)
        }
        else {
            # Un 'dataset' dinámico (que empieza con '@') ya se descartó en el paso 1 (Estado = Dinamico).
            $params.dataset = $fila.SitioDestino
        }

        # Las vistas (parámetro "view") son GUIDs propios de cada lista y no se mapean.
        if (Get-Propiedad $params 'view') {
            Write-Aviso ("{0} '{1}': usa una vista (view = {2}); verificar que exista en el destino." -f $op.Tipo, $op.Ruta, $params.view)
        }

        $cambios.Add([pscustomobject]@{
            Tipo            = $op.Tipo
            Nombre          = $op.Nombre
            Ruta            = $op.Ruta
            Lista           = $fila.NombreLista
            TablaAnterior   = $tablaAnterior
            TablaNueva      = $params.table
            DatasetAnterior = $datasetAnterior
            DatasetNuevo    = $params.dataset
        })
    }

    return , $cambios
}

# Convierte las conexiones Invoker a Embedded, replicando el formato que genera Power Automate
# cuando la conexión es la del propio flujo:
#   - connectionReferences.<ref>.source      -> "Embedded"
#   - inputs.authentication de cada acción   -> "@parameters('$authentication')"
#   - se asegura que exista el parámetro $authentication en la definición
function Convert-InvokerAEmbedded($Raiz, $Invoker) {
    foreach ($r in $Invoker.Referencias) { $r.Referencia.source = 'Embedded' }
    foreach ($a in $Invoker.Acciones)    { $a.Accion.inputs.authentication = "@parameters('`$authentication')" }

    $definicion = Get-NodoDefinicion $Raiz
    $parametros = Get-Propiedad $definicion 'parameters'
    if ($parametros -isnot [System.Management.Automation.PSCustomObject]) {
        $parametros = New-Object PSObject
        $definicion | Add-Member -NotePropertyName 'parameters' -NotePropertyValue $parametros -Force
    }
    if ($null -eq $parametros.PSObject.Properties['$authentication']) {
        $parametros | Add-Member -NotePropertyName '$authentication' -NotePropertyValue ([pscustomobject]@{
            defaultValue = New-Object PSObject
            type         = 'SecureObject'
        })
        Write-Host "  Info: se agregó el parámetro '`$authentication' a la definición." -ForegroundColor DarkGray
    }
}

# Serializa la definición modificada y la guarda en UTF-8 sin BOM.
function Save-NuevaDefinicion($Raiz, [string]$RutaSalida) {
    $json = ConvertTo-Json -InputObject $Raiz -Depth 100
    [System.IO.File]::WriteAllText($RutaSalida, $json, (New-Object System.Text.UTF8Encoding($false)))
}

#endregion

#region Programa principal

try {
    Write-Titulo 'Migración de referencias de SharePoint en definition.json'

    # --- Resolución y verificación de rutas ---
    if ([string]::IsNullOrEmpty($Separador)) { throw 'El separador no puede estar vacío.' }

    $RutaDefinicion = Resolve-RutaCompleta $RutaDefinicion
    if (-not (Test-Path -LiteralPath $RutaDefinicion -PathType Leaf)) {
        throw "No se encontró el archivo de definición: $RutaDefinicion"
    }
    $carpetaDefinicion = Split-Path -Parent $RutaDefinicion

    if ([string]::IsNullOrWhiteSpace($RutaMapeo)) {
        $RutaMapeo = Join-Path $carpetaDefinicion 'list_mapping.csv'
        if (-not (Test-Path -LiteralPath $RutaMapeo -PathType Leaf) -and $PSScriptRoot) {
            $alternativa = Join-Path $PSScriptRoot 'list_mapping.csv'
            if (Test-Path -LiteralPath $alternativa -PathType Leaf) { $RutaMapeo = $alternativa }
        }
    }
    $RutaMapeo = Resolve-RutaCompleta $RutaMapeo
    if (-not (Test-Path -LiteralPath $RutaMapeo -PathType Leaf)) {
        throw "No se encontró el archivo de mapeo: $RutaMapeo"
    }

    if ([string]::IsNullOrWhiteSpace($CarpetaSalida)) { $CarpetaSalida = $carpetaDefinicion }
    $CarpetaSalida = Resolve-RutaCompleta $CarpetaSalida
    $rutaNuevaDefinicion = Join-Path $CarpetaSalida 'new_definition.json'

    Write-Host "  Definición : $RutaDefinicion"
    Write-Host "  Mapeo      : $RutaMapeo (separador '$Separador')"
    Write-Host "  Salida     : $CarpetaSalida"

    # --- Paso 0 ---
    Write-Titulo 'Paso 0: Validación del archivo de mapeo'
    Test-EstructuraMapeo $RutaMapeo $Separador
    $mapeo = Read-Mapeo $RutaMapeo $Separador

    # --- Paso 1 ---
    Write-Titulo 'Paso 1: Relevamiento de GUIDs de listas en definition.json'
    $raiz = Read-Definicion $RutaDefinicion
    $acciones = Get-TodasLasAcciones $raiz
    $relevamiento = Get-RelevamientoGuids $raiz $acciones $mapeo
    $invoker = Find-ConexionesInvoker $raiz $acciones $ConvertirInvokerAEmbedded.IsPresent

    # Crear la carpeta de salida recién ahora (ya validados los datos de entrada)
    if (-not (Test-Path -LiteralPath $CarpetaSalida -PathType Container)) {
        New-Item -ItemType Directory -Path $CarpetaSalida -Force | Out-Null
    }

    # --- Paso 2 ---
    if ($relevamiento.Faltantes.Count -gt 0) {
        Write-Titulo 'Paso 2: GUIDs sin mapear'
        $rutaPendiente = New-MapeoPendiente $RutaMapeo $Separador $relevamiento.Faltantes.ToArray() $CarpetaSalida
        Write-Aviso ("Faltan {0} GUID(s) en el mapeo." -f $relevamiento.Faltantes.Count)
        Write-Aviso "Se generó el archivo: $rutaPendiente"
        Write-Aviso 'Complete nombre_lista, GUID destino y sitio destino, reemplace list_mapping.csv y vuelva a ejecutar.'
        if ($invoker.Hay -and -not $ConvertirInvokerAEmbedded) {
            Write-Aviso 'El flujo además tiene conexiones Invoker: al volver a ejecutar agregue -ConvertirInvokerAEmbedded.'
        }
        if (Test-Path -LiteralPath $rutaNuevaDefinicion) {
            Write-Aviso "Existe un new_definition.json de una ejecución anterior en la carpeta de salida; NO corresponde a esta ejecución."
        }
        Write-Fallo 'Ejecución cancelada: no se generó new_definition.json para evitar una migración parcial.'
        exit 2
    }

    # --- Conexiones Invoker sin autorización para convertir ---
    if ($invoker.Hay -and -not $ConvertirInvokerAEmbedded) {
        Write-Host ''
        Write-Fallo 'Ejecución cancelada: el flujo tiene conexiones Invoker y el paquete no se podría importar.'
        Write-Fallo 'Vuelva a ejecutar con -ConvertirInvokerAEmbedded para convertirlas a Embedded.'
        Write-Fallo 'No se generó new_definition.json.'
        exit 3
    }

    # --- Paso 3 ---
    Write-Titulo 'Paso 3: Generación de new_definition.json'
    $cambios = Update-ReferenciasSharePoint $relevamiento.Operaciones $mapeo
    if ($invoker.Hay) { Convert-InvokerAEmbedded $raiz $invoker }
    Save-NuevaDefinicion $raiz $rutaNuevaDefinicion

    Write-Host ''
    Write-Host ("  Operaciones modificadas: {0}" -f $cambios.Count) -ForegroundColor Cyan
    $n = 0
    foreach ($c in $cambios) {
        $n++
        Write-Host ''
        Write-Host ("  [{0}] {1}: {2}" -f $n, $c.Tipo, $c.Nombre) -ForegroundColor White
        if ($c.Ruta -ne $c.Nombre) { Write-Host "      Ubicación: $($c.Ruta)" }
        Write-Host "      Lista    : $($c.Lista)"
        Write-Host "      table    : $($c.TablaAnterior) -> $($c.TablaNueva)"
        Write-Host "      dataset  : $($c.DatasetAnterior) -> $($c.DatasetNuevo)"
    }
    if ($invoker.Hay) {
        Write-Host ''
        Write-Host '  Conexiones convertidas de Invoker a Embedded:' -ForegroundColor Cyan
        foreach ($r in $invoker.Referencias) { Write-Host "      Referencia de conexión: $($r.Nombre)  (source: Invoker -> Embedded)" }
        foreach ($a in $invoker.Acciones) {
            Write-Host ("      {0}: {1}  (authentication -> @parameters('`$authentication'))" -f $a.Tipo, $a.Ruta)
        }
    }

    Write-Host ''
    Write-Ok "Archivo generado: $rutaNuevaDefinicion"

    # Recordatorio al final para que no se pierda entre el resto de la salida
    if ($invoker.Hay) {
        Write-Host ''
        Write-Host '  *************************** ACCIÓN REQUERIDA DESPUÉS DE IMPORTAR ***************************' -ForegroundColor Yellow
        Write-Aviso 'Las conexiones Invoker se convirtieron a Embedded. Hasta que se restauren, TODAS las acciones'
        Write-Aviso 'se ejecutan con la conexión elegida al importar (sus permisos y su identidad en SharePoint).'
        Write-Aviso "Para restaurar: en el flujo importado, 'Usuarios que solo pueden ejecutar' -> marcar la conexión"
        Write-Aviso "como 'Proporcionada por el usuario que solo tiene permisos de ejecución'."
        foreach ($r in $invoker.Referencias) { Write-Aviso "  Conexión a restaurar: $($r.Nombre)" }
        Write-Host '  *********************************************************************************************' -ForegroundColor Yellow
    }
    exit 0
}
catch {
    Write-Host ''
    Write-Fallo $_.Exception.Message
    Write-Fallo 'Ejecución cancelada.'
    exit 1
}

#endregion
