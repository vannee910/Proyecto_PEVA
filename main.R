# Procesamiento de datos 
# Perez Espinoza Vanessa Anahi
# No. cuenta: 318319272

rm(list = ls())

# Cargar librerías y funciones
library(readxl)
library(dplyr)
library(openxlsx)
source("scripts/objetos.R")
source("scripts/funciones.R")
 

# Crear carpeta de salida si no existe
if(!dir.exists("docs")) dir.create("docs")

# Rutas de archivos
path_cn_raw <- "data/capitalneto_proyecto.xlsx"
path_rc_raw <- "data/riesgo_credito.xlsx"
path_plantilla_icap <- "data/RESUMEN ICAP MODIFICADO.xlsx"

# Verificar que existan los archivos
if(!file.exists(path_cn_raw)) stop("ERROR: No se encuentra capitalneto_proyecto.xlsx")
if(!file.exists(path_rc_raw)) stop("ERROR: No se encuentra riesgo_credito.xlsx")
if(!file.exists(path_plantilla_icap)) stop("ERROR: No se encuentra RESUMEN ICAP MODIFICADO.xlsx")


# PROCESAR CAPITAL NETO (ARCHIVO SEPARADO 1)

cat("Procesando Capital Neto...")
datos_cn_raw <- read_xlsx(path_cn_raw, sheet = 2) %>% 
  mutate(fecha = as.Date(fecha))

# Usar la fecha específica
fecha_analisis <- max(datos_cn_raw$fecha, na.rm = TRUE)
cat("   Fecha de análisis:", as.character(fecha_analisis), "\n")

# Calcular capital neto 
resultados_cn <- calcular_capital_neto(datos_cn_raw, fecha_analisis)

# Guardar archivo INDEPENDIENTE 1: CAPITAL_NETO.xlsx
cat("Guardando CAPITAL_NETO.xlsx...")
wb_cn <- createWorkbook()
addWorksheet(wb_cn, "Capital_Neto")
writeData(wb_cn, sheet = 1, x = resultados_cn)
saveWorkbook(wb_cn, "docs/CAPITAL_NETO.xlsx", overwrite = TRUE)

# PROCESAR RIESGO DE CRÉDITO (ARCHIVO SEPARADO 2)

cat("Procesando Riesgo de Crédito...\n")
datos_rc_raw <- read_xlsx(path_rc_raw, sheet = "BASE")

# Calcular activos ponderados
resultados_rc <- act_pond(datos_rc_raw)

# Guardar archivo INDEPENDIENTE 2: ACTIVOS_RIESGO.xlsx
cat("Guardando ACTIVOS_RIESGO.xlsx...\n")
wb_rc <- createWorkbook()

# Hoja 1: Desglose por ponderador
addWorksheet(wb_rc, "Desglose_Ponderadores")
writeData(wb_rc, sheet = 1, x = resultados_rc$desglose)

# Hoja 2: Totales
addWorksheet(wb_rc, "Totales")
writeData(wb_rc, sheet = 2, x = resultados_rc$totales)

# Hoja 3: Base completa 
addWorksheet(wb_rc, "Base_Completa")
writeData(wb_rc, sheet = 3, x = resultados_rc$base_completa)

saveWorkbook(wb_rc, "docs/ACTIVOS_RIESGO.xlsx", overwrite = TRUE)


#LLENAR PLANTILLA RESUMEN ICAP (ARCHIVO SEPARADO 3)

cat("Llenando plantilla RESUMEN_ICAP.xlsx...")

# Copiar plantilla original
file.copy(path_plantilla_icap, "docs/RESUMEN_ICAP.xlsx", overwrite = TRUE)

# Cargar workbook
wb_icap <- loadWorkbook("docs/RESUMEN_ICAP.xlsx")

# Usar función para integrar los datos
wb_icap <- integrar_icap(
  wb = wb_icap,
  tabla_cn = resultados_cn,
  tabla_riesgo = resultados_rc,
  nombre_hoja = "VISTA RESUMEN"  
)

# Guardar plantilla modificada
saveWorkbook(wb_icap, "docs/RESUMEN_ICAP.xlsx", overwrite = TRUE)

# GENERAR GRÁFICAS Y GUARDARLAS

cat("Generando gráficas...")

# Generar gráficas de capital neto
graficas_cn <- graficas_CN(datos_cn_raw)

# Guardar gráficas como imágenes
png("docs/grafica_evolucion_activos.png", width = 800, height = 600)
print(graficas_cn$evolucion)  
dev.off()

png("docs/grafica_estructura_capital.png", width = 800, height = 600)
print(graficas_cn$estructura)
dev.off()

# Gráficas de riesgo
if(!is.null(resultados_rc$desglose)) {
  graficas_rc <- graficas_riesgo(resultados_rc$desglose)
  
  if(!is.null(graficas_rc$barras)) {
    png("docs/grafica_exposicion_riesgo.png", width = 800, height = 600)
    print(graficas_rc$barras)
    dev.off()
  }
  
  if(!is.null(graficas_rc$pastel)) {
    png("docs/grafica_distribucion_riesgo.png", width = 800, height = 600)
    print(graficas_rc$pastel)
    dev.off()
  }
  cat("Gráficas de Riesgo generadas\n")
}

# RESUMEN FINAL

cat("PROCESO COMPLETADO MILAGROSAMENTE JAJA :)")

cat("ARCHIVOS GENERADOS EN LA CARPETA 'docs/':")
cat("   1. CAPITAL_NETO.xlsx\n")
cat("   2. ACTIVOS_RIESGO.xlsx\n")
cat("   3. RESUMEN_ICAP.xlsx\n")
cat("   4. Gráficas en formato PNG\n\n")

cat("EJECUTA EL DASHBOARD:\n")
cat("   > shiny::runApp('app.R')\n\n")

cat("Fecha de análisis:", as.character(fecha_analisis), "\n")
cat("Total archivos generados: 7\n")

