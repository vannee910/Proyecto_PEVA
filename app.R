# Dashboard Shiny Integrado Final
# Perez Espinoza Vanessa Anahi
# No. cuenta: 318319272
# Proyecto Final: Módulo 3 - ICAP Basilea I

# LIBRERÍAS
library(shiny)
library(shinydashboard)
library(plotly)
library(DT)
library(readxl)
library(dplyr)
library(ggplot2)
library(openxlsx)

# CARGAR FUNCIONES
if(file.exists("scripts/objetos.R")) source("scripts/objetos.R")
if(file.exists("scripts/funciones.R")) source("scripts/funciones.R")

# UI
ui <- dashboardPage(
  
  dashboardHeader(title = "Dashboard ICAP - Basilea III"),
  
  dashboardSidebar(
    sidebarMenu(
      menuItem("Inicio", tabName = "inicio", icon = icon("home")),
      menuItem("Resumen", tabName = "resumen", icon = icon("dashboard")),
      menuItem("Capital Neto", tabName = "capital", icon = icon("bank")),
      menuItem("Riesgo Crédito", tabName = "riesgo", icon = icon("chart-line")),
      menuItem("Reportes", tabName = "reportes", icon = icon("file-excel"))
    )
  ),
  
  dashboardBody(
    tabItems(
      
      # Pestaña 1: Inicio
      tabItem(
        tabName = "inicio",
        h2("Bienvenido al Dashboard ICAP"),
        fluidRow(
          box(
            width = 12,
            status = "primary",
            solidHeader = TRUE,
            title = "Proyecto Final - Módulo 3",
            tags$div(
              style = "padding: 20px;",
              h4("Seminario de Finanzas I - UNAM"),
              p("Introducción al Estándar Internacional de Basilea III"),
              br(),
              p("Alumno: Perez Espinoza Vanessa Anahi"),
              p("No. Cuenta: 318319272"),
              p("Grupo: 9279"),
              br(),
              actionButton("btn_cargar", "Cargar Datos", 
                           icon = icon("upload"),
                           class = "btn-success",
                           style = "width: 100%;")
            )
          )
        )
      ),
      
      # Pestaña 2: Resumen
      tabItem(
        tabName = "resumen",
        h2("Resumen Ejecutivo"),
        fluidRow(
          valueBoxOutput("vb_icap", width = 3),
          valueBoxOutput("vb_capital_neto", width = 3),
          valueBoxOutput("vb_activos_riesgo", width = 3),
          valueBoxOutput("vb_coef_basico", width = 3)
        ),
        fluidRow(
          column(6,
                 box(
                   title = "Composición Capital",
                   width = 12,
                   plotlyOutput("grafica_composicion")
                 )
          ),
          column(6,
                 box(
                   title = "Distribución Riesgo",
                   width = 12,
                   plotlyOutput("grafica_riesgo")
                 )
          )
        )
      ),
      
      # Pestaña 3: Capital Neto
      tabItem(
        tabName = "capital",
        h2("Capital Neto"),
        fluidRow(
          box(
            title = "Detalle del Capital Neto",
            width = 12,
            DTOutput("tabla_cn")
          )
        ),
        fluidRow(
          column(6,
                 box(
                   title = "Evolución Activos",
                   width = 12,
                   plotOutput("grafica_evolucion")
                 )
          ),
          column(6,
                 box(
                   title = "Estructura Capital",
                   width = 12,
                   plotOutput("grafica_estructura")
                 )
          )
        )
      ),
      
      # Pestaña 4: Riesgo Crédito
      tabItem(
        tabName = "riesgo",
        h2("Riesgo de Crédito"),
        fluidRow(
          column(6,
                 box(
                   title = "Desglose por Ponderador",
                   width = 12,
                   DTOutput("tabla_desglose")
                 )
          ),
          column(6,
                 box(
                   title = "Totales",
                   width = 12,
                   DTOutput("tabla_totales")
                 )
          )
        ),
        fluidRow(
          column(6,
                 box(
                   title = "Exposición por Ponderador",
                   width = 12,
                   plotOutput("grafica_exposicion")
                 )
          ),
          column(6,
                 box(
                   title = "Distribución Riesgo",
                   width = 12,
                   plotOutput("grafica_distribucion")
                 )
          )
        )
      ),
      
      # Pestaña 5: Reportes
      tabItem(
        tabName = "reportes",
        h2("Reportes Excel"),
        fluidRow(
          column(4,
                 box(
                   title = "Capital Neto",
                   width = 12,
                   downloadButton("descargar_cn", "Descargar")
                 )
          ),
          column(4,
                 box(
                   title = "Activos Riesgo",
                   width = 12,
                   downloadButton("descargar_ar", "Descargar")
                 )
          ),
          column(4,
                 box(
                   title = "Resumen ICAP",
                   width = 12,
                   downloadButton("descargar_icap", "Descargar")
                 )
          )
        )
      )
    )
  )
)

# SERVER
server <- function(input, output, session) {
  
  # Datos reactivos
  datos <- reactiveValues(
    capital_neto = NULL,
    riesgo = NULL,
    graficas_cn = NULL,
    graficas_rc = NULL
  )
  
  # Botón para procesar - EJECUTA main.R real
  observeEvent(input$btn_cargar, {
    showNotification("Ejecutando main.R...", type = "message", duration = NULL)
    
    tryCatch({
      # Ejecutar main.R
      source("main.R", local = TRUE)
      
      # Cargar los archivos generados
      if(file.exists("docs/CAPITAL_NETO.xlsx")) {
        datos$capital_neto <- read_xlsx("docs/CAPITAL_NETO.xlsx")
      }
      
      if(file.exists("docs/ACTIVOS_RIESGO.xlsx")) {
        datos$riesgo <- list(
          desglose = read_xlsx("docs/ACTIVOS_RIESGO.xlsx", sheet = "Desglose_Ponderadores"),
          totales = read_xlsx("docs/ACTIVOS_RIESGO.xlsx", sheet = "Totales")
        )
      }
      
      # Cargar datos para gráficas de capital neto
      if(file.exists("data/capitalneto_proyecto.xlsx")) {
        datos_raw <- read_xlsx("data/capitalneto_proyecto.xlsx", sheet = 2) %>% 
          mutate(fecha = as.Date(fecha))
        datos$graficas_cn <- graficas_CN(datos_raw)
      }
      
      # Cargar datos para gráficas de riesgo
      if(!is.null(datos$riesgo$desglose)) {
        datos$graficas_rc <- graficas_riesgo(datos$riesgo$desglose)
      }
      
      showNotification("¡Procesamiento completado!", type = "message", duration = 5)
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error", duration = 10)
    })
  })
  
  # Value boxes
  output$vb_icap <- renderValueBox({
    if(is.null(datos$capital_neto)) {
      return(valueBox("N/A", "ICAP", icon = icon("percent"), color = "blue"))
    }
    
    icap_valor <- datos$capital_neto %>% 
      filter(grepl("ÍNDICE DE CAPITALIZACIÓN", Concepto)) %>% 
      pull(Monto)
    
    if(length(icap_valor) == 0) icap_valor <- 0
    
    color <- if(icap_valor >= 0.08) "green" else if(icap_valor >= 0.06) "yellow" else "red"
    
    valueBox(
      paste0(round(icap_valor * 100, 2), "%"),
      "ICAP",
      icon = icon("percent"),
      color = color
    )
  })
  
  output$vb_capital_neto <- renderValueBox({
    if(is.null(datos$capital_neto)) {
      return(valueBox("N/A", "Capital Neto", icon = icon("bank"), color = "blue"))
    }
    
    cn_valor <- datos$capital_neto %>% 
      filter(grepl("CAPITAL NETO", Concepto)) %>% 
      pull(Monto)
    
    if(length(cn_valor) == 0) cn_valor <- 0
    
    valueBox(
      paste0("$", format(round(cn_valor/1e6, 1), big.mark = ","), "M"),
      "Capital Neto",
      icon = icon("bank"),
      color = "blue"
    )
  })
  
  output$vb_activos_riesgo <- renderValueBox({
    if(is.null(datos$riesgo$totales)) {
      return(valueBox("N/A", "Activos Riesgo", icon = icon("exclamation-triangle"), color = "yellow"))
    }
    
    riesgo_valor <- datos$riesgo$totales %>% 
      filter(grepl("Total Activos", Concepto)) %>% 
      pull(Monto)
    
    if(length(riesgo_valor) == 0) riesgo_valor <- 0
    
    valueBox(
      paste0("$", format(round(riesgo_valor/1e6, 1), big.mark = ","), "M"),
      "Activos Ponderados",
      icon = icon("exclamation-triangle"),
      color = "yellow"
    )
  })
  
  output$vb_coef_basico <- renderValueBox({
    if(is.null(datos$capital_neto)) {
      return(valueBox("N/A", "Coef. Básico", icon = icon("chart-bar"), color = "purple"))
    }
    
    coef_valor <- datos$capital_neto %>% 
      filter(grepl("COEFICIENTE DE CAPITAL BÁSICO", Concepto)) %>% 
      pull(Monto)
    
    if(length(coef_valor) == 0) coef_valor <- 0
    
    color <- if(coef_valor >= 0.045) "green" else if(coef_valor >= 0.03) "yellow" else "red"
    
    valueBox(
      paste0(round(coef_valor * 100, 2), "%"),
      "Coef. Capital Básico",
      icon = icon("chart-bar"),
      color = color
    )
  })
  
  # Gráficas Resumen
  output$grafica_composicion <- renderPlotly({
    if(is.null(datos$capital_neto)) return(plotly_empty())
    
    componentes <- datos$capital_neto %>%
      filter(Concepto %in% c(
        "I. CAPITAL CONTRIBUIDO",
        "II. CAPITAL GANADO",
        "III. INVERSIONES Y DEDUCCIONES"
      )) %>%
      mutate(
        Componente = case_when(
          grepl("CONTRIBUIDO", Concepto) ~ "Capital Contribuido",
          grepl("GANADO", Concepto) ~ "Capital Ganado",
          grepl("DEDUCCIONES", Concepto) ~ "Deducciones"
        ),
        Monto_Absoluto = abs(Monto)
      )
    
    if(nrow(componentes) == 0) return(plotly_empty())
    
    plot_ly(
      componentes,
      labels = ~Componente,
      values = ~Monto_Absoluto,
      type = 'pie',
      hole = 0.4,
      textinfo = 'label+percent',
      marker = list(colors = c('#3498db', '#2ecc71', '#e74c3c'))
    ) %>%
      layout(
        title = "Composición del Capital",
        showlegend = TRUE
      )
  })
  
  output$grafica_riesgo <- renderPlotly({
    if(is.null(datos$graficas_rc) || is.null(datos$graficas_rc$barras)) {
      return(plotly_empty() %>%
               layout(title = "Cargue datos para ver la gráfica"))
    }
    
    ggplotly(datos$graficas_rc$barras)
  })
  
  # Tablas Capital Neto
  output$tabla_cn <- renderDT({
    if(is.null(datos$capital_neto)) return(data.frame())
    
    datos$capital_neto %>%
      mutate(
        Valor = ifelse(
          grepl("COEFICIENTE|ÍNDICE", Concepto),
          paste0(round(Monto * 100, 4), "%"),
          paste0("$", format(round(Monto, 2), big.mark = ",", scientific = FALSE))
        )
      ) %>%
      select(Concepto, Valor) %>%
      datatable(
        options = list(pageLength = 25),
        rownames = FALSE
      )
  })
  
  # Gráficas Capital Neto
  output$grafica_evolucion <- renderPlot({
    if(is.null(datos$graficas_cn$evolucion)) {
      return(ggplot() + 
               labs(title = "Cargue datos para ver la gráfica") +
               theme_void())
    }
    datos$graficas_cn$evolucion
  })
  
  output$grafica_estructura <- renderPlot({
    if(is.null(datos$graficas_cn$estructura)) {
      return(ggplot() + 
               labs(title = "Cargue datos para ver la gráfica") +
               theme_void())
    }
    datos$graficas_cn$estructura
  })
  
  # Tablas Riesgo
  output$tabla_desglose <- renderDT({
    if(is.null(datos$riesgo$desglose)) return(data.frame())
    
    datos$riesgo$desglose %>%
      mutate(
        Ponderador = paste0(round(Ponderador_Num * 100, 0), "%"),
        Monto = paste0("$", format(round(Monto_Total, 2), big.mark = ",")),
        `Capital Req.` = paste0("$", format(round(Capital_Req_Total, 2), big.mark = ","))
      ) %>%
      select(Ponderador, Monto, `Capital Req.`) %>%
      datatable(
        options = list(pageLength = 10),
        rownames = FALSE
      )
  })
  
  output$tabla_totales <- renderDT({
    if(is.null(datos$riesgo$totales)) return(data.frame())
    
    datos$riesgo$totales %>%
      mutate(
        Valor = paste0("$", format(round(Monto, 2), big.mark = ","))
      ) %>%
      select(Concepto, Valor) %>%
      datatable(
        options = list(pageLength = 10),
        rownames = FALSE
      )
  })
  
  # Gráficas Riesgo
  output$grafica_exposicion <- renderPlot({
    if(is.null(datos$graficas_rc$barras)) {
      return(ggplot() + 
               labs(title = "Cargue datos para ver la gráfica") +
               theme_void())
    }
    datos$graficas_rc$barras
  })
  
  output$grafica_distribucion <- renderPlot({
    if(is.null(datos$graficas_rc$pastel)) {
      return(ggplot() + 
               labs(title = "Cargue datos para ver la gráfica") +
               theme_void())
    }
    datos$graficas_rc$pastel
  })
  
  # Descargas
  output$descargar_cn <- downloadHandler(
    filename = function() { "CAPITAL_NETO.xlsx" },
    content = function(file) {
      file.copy("docs/CAPITAL_NETO.xlsx", file)
    }
  )
  
  output$descargar_ar <- downloadHandler(
    filename = function() { "ACTIVOS_RIESGO.xlsx" },
    content = function(file) {
      file.copy("docs/ACTIVOS_RIESGO.xlsx", file)
    }
  )
  
  output$descargar_icap <- downloadHandler(
    filename = function() { "RESUMEN_ICAP.xlsx" },
    content = function(file) {
      file.copy("docs/RESUMEN_ICAP.xlsx", file)
    }
  )
}

# EJECUTAR APLICACIÓN
shinyApp(ui = ui, server = server)