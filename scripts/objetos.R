
library(readxl)
library(dplyr)
library(tidyr)
library(tibble)
library(ggplot2)
library(openxlsx)
library(stringr)

# Plantilla Base para CAPITAL NETO 
plantilla_cn_base <- tribble(
  ~Concepto,                                                          ~Codigo_Banxico,
  "ACTIVOS SUJETOS A RIESGO TOTALES",                                 90001,
  "ACTIVOS SUJETOS A RIESGO DE CRÉDITO (MÉTODO ESTÁNDAR)",            90002,
  "ACTIVOS SUJETOS A RIESGO DE CRÉDITO (MÉTODO INTERNO)",             90003,
  
  "I. CAPITAL CONTRIBUIDO",                                           NA,
  "  (+) Títulos representativos de capital social",                  95005,
  "  (+) Prima en venta de acciones",                                 95010,
  "  (+) Aportaciones para futuros aumentos",                         95035,
  
  "II. CAPITAL GANADO",                                               NA,
  "  (+) Reservas de capital",                                        95045,
  "  (+) Resultado de Ejercicios Anteriores",                         95050,
  "  (+) Resultado Neto",                                             95055,
  "  (+) Resultado por valuación",                                    95060,
  "  (+) Resultado remediciones beneficios empleados",                95068,
  
  "III. INVERSIONES Y DEDUCCIONES",                                   NA,
  "  (-) Inversiones en otras entidades financieras",                 95150, 
  "  (-) Activos Intangibles",                                        95315,
  "  (-) Impuestos Diferidos",                                        95380,
  "  (-) Operaciones con Personas Relacionadas",                      93916,
  
  "CAPITAL FUNDAMENTAL",                                              NA,
  
  "PARTE BÁSICA NO FUNDAMENTAL",                                      NA,
  "  (+) Instrumentos de capital (Anexo 1-R)",                       95075,
  "  (+) Obligaciones Subordinadas (Art. 64 Transitorios)",           95550,
  
  "CAPITAL BÁSICO",                                                   NA,
  
  "PARTE COMPLEMENTARIA",                                             NA,
  "  (+) Instrumentos de capital (Anexo 1-S)",                       95080,
  "  (+) Obligaciones Subordinadas (Art. 64 Transitorios)",           95555,
  
  "RESERVAS ADMISIBLES",                                              NA,
  "  (Info) Reservas Admisibles (Método Interno)",                    95691,
  "  (Info) Pérdidas Esperadas (Método Interno)",                     95695,
  "  (+) Reservas Computables (Método Interno - Tope 0.6%)",          NA,
  "  (Info) Reservas Admisibles (Método Estándar)",                   95704,
  "  (Info) Pérdidas Esperadas (MÉtodo Estándar)",                    95708,
  "  (+) Reservas Computables (Método Estándar - Tope 1.25%)",        NA,
  
  "CAPITAL COMPLEMENTARIO",                                           NA,
  "CAPITAL NETO",                                                     NA,
  
  "COEFICIENTE DE CAPITAL FUNDAMENTAL",                               NA,
  "COEFICIENTE DE CAPITAL BÁSICO",                                    NA,
  "ÍNDICE DE CAPITALIZACIÓN (ICAP)",                                  NA
)

# Mapeo para RIESGO DE CRÉDITO 
mapa_ponderadores <- tibble(
  Ponderador_Texto = c("0 %", "10 %", "20 %", "50 %", "100 %", "115 %", "150 %"),
  Ponderador_Num   = c(0, 0.10, 0.20, 0.50, 1.00, 1.15, 1.50)
)

# Etiquetas para búsqueda en RESUMEN ICAP 
etiquetas_icap <- list(
  cap_neto = "CAPITAL NETO",
  riesgo_cred = "Riesgo de Crédito",
  riesgo_merc = "Riesgo de Mercado",
  riesgo_oper = "Riesgo Operacional",
  riesgo_tot  = "Riesgo Total",
  icap_final  = "ICAP"
)