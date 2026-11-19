# Reporte Diario Dólar SET-FX (USD/COP)

Automatización diaria que consulta el mercado SET-FX (`dolar.icap.com.co`), genera el reporte diario en Excel con la gráfica del día, actualiza el consolidado mensual en OneDrive y envía el reporte por correo electrónico.

---

## ¿Qué genera cada ejecución?

1. **Carpeta diaria local** (`Reportes/<Año>/<Mes>/<YYYY-MM-DD>/`):
   - `Dolar-YYYY-MM-DD.csv`: datos crudos intradía.
   - `Grafica-Dolar-YYYY-MM-DD.png`: captura del dashboard SET-ICAP del día.
   - `Reporte-Dolar-YYYY-MM-DD.xlsx`: Excel diario con resumen, gráfica e intervalos de 5, 10, 15, 30 y 60 minutos.
2. **Excel mensual local y en OneDrive** (`Reportes/<Año>/<Mes> <Año>.xlsx`):
   - Acumula una hoja por cada día del mes (`YYYY-MM-DD`). Si se ejecuta más de una vez el mismo día, actualiza la hoja sin duplicarla.
   - Al cambiar de mes o año (ej. diciembre o `2027`), crea automáticamente las carpetas y archivos nuevos.
3. **Correo electrónico diario**:
   - Envía el Excel diario adjunto con el resumen de Apertura, Máximo, Mínimo y Cierre, y un botón directo a la carpeta de OneDrive.

---

## Configuración (`.env`)

El archivo `.env` en la raíz del proyecto debe contener:

```env
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
ONEDRIVE_USER_EMAIL=andres.ardila@phc.com.co
ONEDRIVE_BASE_FOLDER=Dolar
MAIL_FROM=andres.ardila@phc.com.co
MAIL_TO=hernan.corredor@phc.com.co
```

> Para cambiar o agregar destinatarios del correo, edita `MAIL_TO` (separando múltiples correos con coma).

---

## Ejecución manual

```powershell
python Dolar.py
```

---

## Tarea Programada en Windows (Cronjob a las 3:10 p. m.)

### 1. Crear o actualizar la tarea programada (lunes a viernes a las 3:10 p. m.)

Registra la tarea en el Programador de tareas de Windows para que `Dolar.py` se ejecute solo en segundo plano. Si el PC estaba suspendido a las 3:10 p. m., se ejecutará apenas se encienda.

```powershell
$pythonPath = (Get-Command python).Source
$carpeta = "c:\Users\Andres Ardila\Documents\OneDrive_2026-11-19\Script Dolar"

$action = New-ScheduledTaskAction -Execute $pythonPath -Argument "`"$carpeta\Dolar.py`"" -WorkingDirectory $carpeta
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 3:10PM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName "ReporteDolarDiario" -Action $action -Trigger $trigger -Settings $settings -Force
```

### 2. Probar la tarea inmediatamente

Ejecuta la tarea programada en este momento sin esperar a las 3:10 p. m., para confirmar que funciona de forma automática.

```powershell
Start-ScheduledTask -TaskName "ReporteDolarDiario"
```

### 3. Ver estado, última y próxima ejecución

Muestra cuándo fue la última vez que corrió, si terminó sin errores (`LastTaskResult: 0` significa éxito) y la fecha/hora de la próxima ejecución.

```powershell
Get-ScheduledTaskInfo -TaskName "ReporteDolarDiario"
```

### 4. Eliminar la tarea programada

Borra la tarea del Programador de tareas de Windows si ya no deseas que se ejecute automáticamente.

```powershell
Unregister-ScheduledTask -TaskName "ReporteDolarDiario" -Confirm:$false
```