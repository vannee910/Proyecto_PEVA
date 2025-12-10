
library(dplyr)
library(tidyr)
library(ggplot2)
library(openxlsx)
library(stringr)
library(tibble)
library(readxl)
library(scales)
library(plotly)


# CALCULAR CAPITAL NETO

calcular_capital_neto <- function(datos_banco, fecha_analisis) {
  
  # Filtrar datos por fecha
  datos_filtrados <- datos_banco %>% filter(fecha == fecha_analisis)
  
  # Función auxiliar para buscar montos
  buscar_monto <- function(codigo) {
    if (is.na(codigo)) return(0)
    val <- datos_filtrados %>% 
      filter(concepto == codigo) %>% 
      summarise(suma = sum(importe1, na.rm = TRUE)) %>% 
      pull(suma)
    if (length(val) == 0) return(0) else return(val)
  }
  
  # Cargar plantilla base
  reporte <- plantilla_cn_base %>%
    mutate(Monto = sapply(Codigo_Banxico, buscar_monto))
  
  # Función para obtener valores de múltiples códigos
  get_val <- function(cod) { 
    sum(reporte$Monto[reporte$Codigo_Banxico %in% cod], na.rm = TRUE) 
  }
  
  # --- Variables de Insumo ---
  asrc_totales  <- get_val(90001)
  asrc_estandar <- get_val(90002)
  asrc_interno  <- get_val(90003)
  
  # --- Cálculo Reservas ---
  res_adm_int <- get_val(95691)
  perd_esp_int <- get_val(95695)
  exceso_int <- max(0, res_adm_int - perd_esp_int)
  limite_int <- 0.006 * asrc_interno 
  reservas_computables_int <- min(exceso_int, limite_int)
  
  res_adm_est <- get_val(95704)
  perd_esp_est <- get_val(95708)
  exceso_est <- max(0, res_adm_est - perd_esp_est)
  limite_est <- 0.0125 * asrc_estandar
  reservas_computables_est <- min(exceso_est, limite_est)
  
  total_reservas_computables <- reservas_computables_int + reservas_computables_est
  
  # --- Capitales ---
  c_contribuido <- get_val(c(95005, 95010, 95035))
  c_ganado      <- get_val(c(95045, 95050, 95055, 95060, 95068))
  deducciones   <- get_val(c(95150, 95315, 95380, 93916))
  
  cap_fundamental <- c_contribuido + c_ganado - deducciones
  cap_basico_no_fund <- get_val(c(95075, 95550))
  cap_basico <- cap_fundamental + cap_basico_no_fund
  
  cap_complementario_instrumentos <- get_val(c(95080, 95555))
  cap_complementario <- cap_complementario_instrumentos + total_reservas_computables
  cap_neto <- cap_basico + cap_complementario
  
  # --- Índices ---
  coef_fundamental <- if(asrc_totales > 0) cap_fundamental / asrc_totales else 0
  coef_basico      <- if(asrc_totales > 0) cap_basico / asrc_totales else 0
  icap             <- if(asrc_totales > 0) cap_neto / asrc_totales else 0
  
  # --- Inyección Final en Reporte ---
  reporte_final <- reporte %>%
    mutate(Monto = case_when(
      Concepto == "I. CAPITAL CONTRIBUIDO" ~ c_contribuido,
      Concepto == "II. CAPITAL GANADO" ~ c_ganado,
      Concepto == "III. INVERSIONES Y DEDUCCIONES" ~ deducciones,
      Concepto == "CAPITAL FUNDAMENTAL" ~ cap_fundamental,
      Concepto == "PARTE BÁSICA NO FUNDAMENTAL" ~ cap_basico_no_fund,
      Concepto == "CAPITAL BÁSICO" ~ cap_basico,
      Concepto == "PARTE COMPLEMENTARIA" ~ cap_complementario,
      Concepto == "CAPITAL COMPLEMENTARIO" ~ cap_complementario,
      Concepto == "CAPITAL NETO" ~ cap_neto,
      grepl("Reservas Computables.*Interno", Concepto) ~ reservas_computables_int,
      grepl("Reservas Computables.*Estándar", Concepto) ~ reservas_computables_est,
      Concepto == "COEFICIENTE DE CAPITAL FUNDAMENTAL" ~ coef_fundamental,
      Concepto == "COEFICIENTE DE CAPITAL BÁSICO" ~ coef_basico,
      Concepto == "ÍNDICE DE CAPITALIZACIÓN (ICAP)" ~ icap,
      TRUE ~ Monto 
    )) %>%
    select(Concepto, Monto)
  
  return(reporte_final)
}


# CÁLCULO ACTIVOS PONDERADOS (RIESGO CRÉDITO)

act_pond <- function(datos_credito_base) {
  
  # Normalizar nombres de columnas (a minúsculas sin espacios)
  names(datos_credito_base) <- tolower(names(datos_credito_base))
  names(datos_credito_base) <- gsub(" ", "_", names(datos_credito_base))
  
  cat("  - Columnas disponibles después de normalizar:\n")
  cat("    ", paste(names(datos_credito_base), collapse = ", "), "\n")
  
  # Buscar columna de monto - varias posibles opciones
  col_monto <- NULL
  posibles_montos <- c(
    "monto_dispuesto_por_el_acreditado",
    "monto_no_dispuesto",
    "dat_expos_incump_total",
    "exposición_al_incumplimiento_de_la_parte_dispuest",
    "importe",
    "saldo"
  )
  
  for (col in posibles_montos) {
    if (col %in% names(datos_credito_base)) {
      col_monto <- col
      cat("  - Usando columna de monto:", col_monto, "\n")
      break
    }
  }
  
  # Buscar columna de ponderador
  col_pond <- NULL
  posibles_ponds <- c(
    "ponderador_de_riesgo",
    "ponderador",
    "factor_de_riesgo"
  )
  
  for (col in posibles_ponds) {
    if (col %in% names(datos_credito_base)) {
      col_pond <- col
      cat("  - Usando columna de ponderador:", col_pond, "\n")
      break
    }
  }
  
  # Si no encontramos las columnas, mostrar error informativo
  if (is.null(col_monto)) {
    cat("  ❌ ERROR: No se encontró columna de monto. Columnas disponibles:\n")
    cat("    ", paste(names(datos_credito_base), collapse = "\n    "), "\n")
    stop("No se encontró columna de monto en los datos")
  }
  
  if (is.null(col_pond)) {
    cat("  ❌ ERROR: No se encontró columna de ponderador. Columnas disponibles:\n")
    cat("    ", paste(names(datos_credito_base), collapse = "\n    "), "\n")
    stop("No se encontró columna de ponderador en los datos")
  }
  
  # Preparar datos
  df <- datos_credito_base %>%
    rename(monto = all_of(col_monto), 
           pond_raw = all_of(col_pond))
  
  cat("  - Resumen de monto (primeras filas):\n")
  print(head(df$monto))
  cat("  - Resumen de ponderador (primeras filas):\n")
  print(head(df$pond_raw))
  
  # Verificar si hay valores NA en monto
  cat("  - Valores NA en monto:", sum(is.na(df$monto)), "de", nrow(df), "\n")
  cat("  - Valores NA en ponderador:", sum(is.na(df$pond_raw)), "de", nrow(df), "\n")
  
  # Limpiar datos - convertir ponderadores
  limpiar_pct <- function(x) {
    if (is.numeric(x)) {
      # Si ya es numérico, devolver tal cual
      return(as.numeric(x))
    } else if (is.character(x)) {
      # Si es texto, quitar % y convertir a número
      val <- as.numeric(gsub("%", "", x))
      # Si el valor es > 1, probablemente es porcentaje (ej: 50% = 50)
      if (!is.na(val) && val > 1) {
        return(val / 100)
      } else {
        return(val)
      }
    } else {
      return(0)
    }
  }
  
  # Cálculos principales con manejo de NAs
  df_calc <- df %>%
    mutate(
      # Reemplazar NAs en monto con 0
      monto = replace_na(monto, 0),
      
      # Convertir ponderador
      Ponderador_Num = sapply(pond_raw, limpiar_pct),
      Ponderador_Num = replace_na(Ponderador_Num, 0),
      
      # Calcular exposición (en tus datos no parece haber interés)
      exposicion = monto,
      
      # Calcular activos ponderados
      activos_pond = exposicion * Ponderador_Num,
      
      # Calcular requerimiento de capital (8%)
      req_capital = activos_pond * 0.08
    )
  
  # Resumen de cálculos
  cat("  - Resumen de cálculos:\n")
  cat("    Suma de montos:", sum(df_calc$monto, na.rm = TRUE), "\n")
  cat("    Suma de exposición:", sum(df_calc$exposicion, na.rm = TRUE), "\n")
  cat("    Suma de activos_pond:", sum(df_calc$activos_pond, na.rm = TRUE), "\n")
  
  # Tabla Desglose por ponderador
  cat("  - Creando desglose por ponderador...\n")
  tabla_desglose <- df_calc %>%
    group_by(Ponderador_Num) %>%
    summarise(
      Monto_Total = sum(exposicion, na.rm = TRUE),
      Capital_Req_Total = sum(req_capital, na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    ungroup() %>%
    # Ordenar por ponderador
    arrange(Ponderador_Num)
  
  cat("  - Desglose creado con", nrow(tabla_desglose), "filas\n")
  print(tabla_desglose)
  
  # Tabla Totales
  tabla_totales <- tibble(
    Concepto = c("Total Activos Ponderados", "Total Requerimiento Capital"),
    Monto = c(
      sum(df_calc$activos_pond, na.rm = TRUE),
      sum(df_calc$req_capital, na.rm = TRUE)
    )
  )
  
  cat("  - Totales calculados:\n")
  print(tabla_totales)
  
  # Retornar lista completa
  return(list(
    desglose = tabla_desglose,
    totales = tabla_totales,
    base_completa = df_calc
  ))
}


# FUNCIONES PARA GRÁFICAS DE CAPITAL NETO

graficas_CN <- function(datos_banco) {
  
  # Filtrar y preparar datos para evolución
  datos_plot <- datos_banco %>% 
    mutate(fecha = as.Date(fecha)) %>% 
    arrange(fecha)
  
  # Gráfica 1: Evolución de Activos Sujetos a Riesgo
  g1 <- ggplot(filter(datos_plot, concepto == 90001), 
               aes(x = fecha, y = importe1)) +
    geom_line(color = "#2874A6", linewidth = 1.2) + 
    geom_point(color = "#1B4F72", size = 2) +
    geom_area(fill = "#2874A6", alpha = 0.1) +
    labs(
      title = "Evolución de Activos Sujetos a Riesgo", 
      subtitle = "Histórico de exposición total",
      y = "Monto (MXN)", 
      x = "Fecha"
    ) + 
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 16),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  # Gráfica 2: Estructura de Capital
  datos_comp <- filter(datos_plot, concepto %in% c(95005, 95050)) %>% 
    mutate(
      Tipo = case_when(
        concepto == 95005 ~ "Capital Social",
        concepto == 95050 ~ "Resultados Acumulados"
      ),
      fecha = as.Date(fecha)
    )
  
  g2 <- ggplot(datos_comp, aes(x = as.factor(format(fecha, "%Y-%m")), 
                               y = importe1, fill = Tipo)) +
    geom_col(position = "dodge", width = 0.7) +
    scale_fill_manual(values = c("#AED6F1", "#2E86C1")) +
    labs(
      title = "Estructura de Capital", 
      subtitle = "Comparativa: Capital Social vs Resultados Acumulados",
      x = "Fecha", 
      y = "Monto (MXN)"
    ) + 
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 16),
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "bottom"
    )
  
  return(list(evolucion = g1, estructura = g2))
}


# FUNCIONES PARA GRÁFICAS DE RIESGO DE CRÉDITO
graficas_riesgo <- function(tabla_desglose) {
  
  # Verificar que tabla_desglose existe y tiene datos
  if(is.null(tabla_desglose)) {
    message("ERROR: tabla_desglose es NULL")
    return(NULL)
  }
  
  if(nrow(tabla_desglose) == 0) {
    message("ERROR: tabla_desglose tiene 0 filas")
    return(NULL)
  }
  
  # Verificar columnas requeridas
  columnas_requeridas <- c("Ponderador_Num", "Monto_Total")
  if(!all(columnas_requeridas %in% names(tabla_desglose))) {
    message("ERROR: Faltan columnas necesarias: ", 
            paste(columnas_requeridas[!columnas_requeridas %in% names(tabla_desglose)], collapse = ", "))
    return(NULL)
  }
  
  tryCatch({
    # Preparar datos
    datos <- tabla_desglose %>%
      mutate(
        Ponderador_Texto = paste0(round(Ponderador_Num * 100, 0), "%"),
        Porcentaje_Exposicion = ifelse(
          sum(Monto_Total, na.rm = TRUE) > 0,
          Monto_Total / sum(Monto_Total, na.rm = TRUE),
          0
        )
      ) %>%
      arrange(desc(Monto_Total))
    
    # Verificar que hay datos válidos
    if(all(datos$Monto_Total == 0)) {
      message("ADVERTENCIA: Todos los montos son 0")
      return(NULL)
    }
    
    # Gráfica 1: Barras de exposición por ponderador
    g1 <- ggplot(datos, aes(x = reorder(Ponderador_Texto, Ponderador_Num), 
                            y = Monto_Total)) +
      geom_bar(stat = "identity", fill = "#E74C3C", width = 0.7) +
      labs(
        title = "Exposición Total por Ponderador de Riesgo",
        subtitle = "Distribución de cartera según nivel de riesgo",
        x = "Ponderador de Riesgo",
        y = "Exposición Total (MXN)"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 16, hjust = 0.5),
        plot.subtitle = element_text(hjust = 0.5, color = "#7F8C8D"),
        axis.text.x = element_text(angle = 45, hjust = 1, size = 11),
        axis.title = element_text(face = "bold"),
        panel.grid.major.x = element_blank()
      )
    
    # Solo agregar texto si los valores no son muy pequeños
    if(max(datos$Monto_Total) > 1000) {
      g1 <- g1 + 
        geom_text(aes(label = paste0("$", format(round(Monto_Total/1e6, 1), 
                                                 big.mark = ","), "M")),
                  vjust = -0.5, size = 3.5, fontface = "bold")
    }
    
    # Gráfica 2: Pastel de distribución (solo si hay datos positivos)
    datos_pie <- datos %>%
      filter(Monto_Total > 0)
    
    if(nrow(datos_pie) > 0) {
      g2 <- ggplot(datos_pie, aes(x = "", y = Monto_Total, fill = Ponderador_Texto)) +
        geom_bar(stat = "identity", width = 1) +
        coord_polar("y", start = 0) +
        scale_fill_brewer(palette = "Set3") +
        labs(
          title = "Distribución de Exposición por Nivel de Riesgo",
          subtitle = "Porcentaje de la cartera total",
          fill = "Ponderador"
        ) +
        theme_void() +
        theme(
          plot.title = element_text(face = "bold", size = 16, hjust = 0.5),
          plot.subtitle = element_text(hjust = 0.5, color = "#7F8C8D"),
          legend.position = "right",
          legend.title = element_text(face = "bold")
        )
      
      # Solo agregar etiquetas si no son demasiadas
      if(nrow(datos_pie) <= 10) {
        g2 <- g2 + 
          geom_text(aes(label = paste0(round(Porcentaje_Exposicion * 100, 1), "%")), 
                    position = position_stack(vjust = 0.5),
                    size = 3.5, fontface = "bold")
      }
    } else {
      g2 <- NULL
    }
    
    return(list(barras = g1, pastel = g2))
    
  }, error = function(e) {
    message("ERROR en graficas_riesgo: ", e$message)
    return(NULL)
  })
}



# FUNCIONES PARA LLENAR PLANTILLAS EXCEL


# Función para volcar datos de capital neto en plantilla
volcar_cn_plantilla <- function(wb, datos_cn, nombre_hoja = "Capital_Neto") {
  if(!nombre_hoja %in% names(wb)) {
    addWorksheet(wb, nombre_hoja)
  }
  writeData(wb, sheet = nombre_hoja, x = datos_cn)
  
  # Aplicar formato de moneda
  addStyle(
    wb, 
    sheet = nombre_hoja, 
    style = createStyle(numFmt = "#,##0.00"), 
    rows = 2:(nrow(datos_cn) + 1), 
    cols = 2
  )
  
  return(wb)
}

# Función para reporte de ponderadores
reporte_pond <- function(wb, resultados_pond, nombre_hoja = "REPORTE") {
  
  if(!nombre_hoja %in% names(wb)) {
    addWorksheet(wb, nombre_hoja)
  }
  
  # Crear tabla visual con mapeo de ponderadores
  tabla_visual <- mapa_ponderadores %>%
    left_join(resultados_pond$desglose, by = "Ponderador_Num") %>%
    mutate(
      Importe = replace_na(Monto_Total, 0),
      Capital_Neto_Requerido = replace_na(Capital_Req_Total, 0)
    ) %>%
    select(Ponderador_Texto, Importe, Capital_Neto_Requerido)
  
  # Escribir desglose
  writeData(wb, sheet = nombre_hoja, x = tabla_visual, 
            startRow = 7, startCol = 2, colNames = FALSE)
  
  # Escribir totales
  fila_totales <- 7 + nrow(tabla_visual) + 1
  writeData(wb, sheet = nombre_hoja, x = resultados_pond$totales, 
            startRow = fila_totales, startCol = 2, colNames = FALSE)
  
  return(wb)
}

# Función para integrar ICAP (llenar plantilla RESUMEN)
integrar_icap <- function(wb, tabla_cn, tabla_riesgo, nombre_hoja = "VISTA RESUMEN") {
  
  # 1. Extraer CAPITAL NETO
  val_cn <- tabla_cn %>% 
    filter(Concepto == "CAPITAL NETO") %>% 
    pull(Monto)
  if(length(val_cn) == 0 || is.na(val_cn)) val_cn <- 0
  
  # 2. Extraer RIESGO CRÉDITO
  val_ap_credito <- tabla_riesgo$totales$Monto[1]
  if(is.na(val_ap_credito)) val_ap_credito <- 0
  
  # 3. Leer hoja existente
  datos_actuales <- tryCatch({
    readWorkbook(wb, sheet = nombre_hoja, colNames = FALSE, skipEmptyRows = FALSE)
  }, error = function(e) {
    stop(paste("Error al leer hoja:", e$message))
  })
  
  # Función auxiliar para buscar filas
  buscar_fila <- function(texto) {
    idx <- grep(texto, datos_actuales[[2]], ignore.case = TRUE)
    if(length(idx) == 0) return(NA)
    return(idx[1])
  }
  
  # Identificar filas
  fila_cn    <- buscar_fila("CAPITAL NETO")
  fila_cred  <- buscar_fila("Riesgo de Crédito")
  fila_merc  <- buscar_fila("Riesgo de Mercado")
  fila_oper  <- buscar_fila("Riesgo Operacional")
  fila_tot   <- buscar_fila("Riesgo Total|Activos Ponderados Totales")
  fila_icap  <- buscar_fila("ICAP|Índice de Capitalización")
  
  # Función para leer valores existentes
  leer_val_excel <- function(fila) {
    if(is.na(fila)) return(0)
    val <- datos_actuales[fila, 3]
    if(is.na(val) || val == "" || val == "NA" || val == "-") return(0)
    num <- suppressWarnings(as.numeric(val))
    if(is.na(num)) return(0)
    return(num)
  }
  
  # Leer valores existentes de mercado y operacional
  val_ap_mercado <- leer_val_excel(fila_merc)
  val_ap_operacional <- leer_val_excel(fila_oper)
  
  # 4. Cálculos finales
  val_ap_totales <- val_ap_credito + val_ap_mercado + val_ap_operacional
  if(is.na(val_ap_totales)) val_ap_totales <- 0
  
  val_icap <- if(val_ap_totales > 0) val_cn / val_ap_totales else 0
  
  # 5. Escribir en Excel
  if(!is.na(fila_cn)) {
    writeData(wb, sheet = nombre_hoja, x = val_cn, 
              startRow = fila_cn, startCol = 3, colNames = FALSE)
  }
  
  if(!is.na(fila_cred)) {
    writeData(wb, sheet = nombre_hoja, x = val_ap_credito, 
              startRow = fila_cred, startCol = 3, colNames = FALSE)
  }
  
  if(!is.na(fila_tot)) {
    writeData(wb, sheet = nombre_hoja, x = val_ap_totales, 
              startRow = fila_tot, startCol = 3, colNames = FALSE)
  }
  
  if(!is.na(fila_icap)) {
    writeData(wb, sheet = nombre_hoja, x = val_icap, 
              startRow = fila_icap, startCol = 3, colNames = FALSE)
    addStyle(wb, sheet = nombre_hoja, 
             style = createStyle(numFmt = "0.00%"), 
             rows = fila_icap, cols = 3)
  }
  
  # Aplicar formato de moneda
  filas_montos <- c(fila_cn, fila_cred, fila_tot)
  filas_montos <- filas_montos[!is.na(filas_montos)]
  
  if(length(filas_montos) > 0) {
    addStyle(wb, sheet = nombre_hoja, 
             style = createStyle(numFmt = "#,##0.00"), 
             rows = filas_montos, cols = 3)
  }
  
  return(wb)
}


# FUNCIÓN PARA EXPORTAR A EXCEL (ARCHIVOS SEPARADOS)

exportar_excel <- function(datos, nombre_archivo, nombre_hoja = "Datos") {
  wb <- createWorkbook()
  addWorksheet(wb, nombre_hoja)
  writeData(wb, sheet = nombre_hoja, x = datos)
  
  # Aplicar formato automático
  if(ncol(datos) >= 2 && is.numeric(datos[[2]])) {
    addStyle(
      wb, 
      sheet = nombre_hoja, 
      style = createStyle(numFmt = "#,##0.00"), 
      rows = 2:(nrow(datos) + 1), 
      cols = 2
    )
  }
  
  saveWorkbook(wb, nombre_archivo, overwrite = TRUE)
  return(TRUE)
}

