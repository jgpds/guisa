# Gerenciamento de pacotes
if (!require("pacman")) install.packages("pacman")
pacman::p_load(
  shiny,
  bslib,
  shinyWidgets,
  DT,
  rbcb,
  lubridate,
  openxlsx,
  readxl,
  dplyr
)

# Constantes e configurações
CONFIG <- list(
  THEME = bs_theme(
    bg = "#ffffff",
    fg = "#333333",
    primary = "#575152",
    base_font = "Montserrat",
    heading_font = "Montserrat",
    font_scale = 1.1
  ),
  INDEXADORES = c("IPCA" = 433, "IGPM" = 189, "INPC" = 188, "SELIC" = 432, "TR" = 226),
  DATE_MIN = "2000-01",
  DATE_MAX = format(Sys.Date(), "%Y-%m")
)

# Funções auxiliares
get_accumulated_index <- function(index_code, start_date, end_date) {
  tryCatch({
    serie <- rbcb::get_series(
      index_code, 
      start_date = start_date, 
      end_date = end_date
    )
    accumulated <- prod(1 + (serie[2] / 100)) - 1
    return(list(success = TRUE, value = accumulated, error = NULL))
  }, error = function(e) {
    return(list(success = FALSE, value = NULL, error = e$message))
  })
}

# UI
ui <- navbarPage(
  theme = CONFIG$THEME,
  
  tags$head(
    includeCSS("www/styles.css"),
    tags$style(HTML("
      .navbar {
        padding-top: 0px;         /* Reduz o padding superior */
        padding-bottom: 0px;      /* Reduz o padding inferior */
      }
    "))
    
  ),
  title = tags$div(
    tags$a(
      href = "#",  # Link para onde o logo leva, se aplicável
      tags$img(src = "Screenshot_2.png", height = "75px", style = "margin-right: 15px;")  # Caminho da imagem da logo
    )
  ),
  # Tab Principal
  tabPanel("Principal",
           tags$head(
             includeCSS("www/styles.css")
           ),
           
           sidebarLayout(
             sidebarPanel(
               selectInput(
                 "indexador",
                 tags$span(icon("chart-line"), "Indexador Econômico"),
                 choices = names(CONFIG$INDEXADORES),
                 selected = "IPCA"
               ),
               
               airMonthpickerInput(
                 "data_inicio",
                 tags$span(icon("calendar"), "Data de Início"),
                 minDate = CONFIG$DATE_MIN,
                 maxDate = CONFIG$DATE_MAX,
                 view = "months",
                 dateFormat = "MM/yyyy"
               ),
               
               airMonthpickerInput(
                 "data_fim",
                 tags$span(icon("calendar-check"), "Data de Término"),
                 minDate = CONFIG$DATE_MIN,
                 maxDate = CONFIG$DATE_MAX,
                 view = "months",
                 dateFormat = "MM/yyyy"
               ),
               
               selectizeInput(
                 "margem_tolerancia",
                 tags$span(icon("sliders-h"), "Margem de Tolerância (%):"),
                 choices = NULL,
                 multiple = TRUE
               ),
               
               hr(),
               
               fileInput(
                 "base1",
                 tags$span(icon("file-excel"), "Base 1 (Mais antiga)"),
                 accept = c(".xlsx", ".xls")
               ),
               
               fileInput(
                 "base2",
                 tags$span(icon("file-excel"), "Base 2 (Mais recente)"),
                 accept = c(".xlsx", ".xls")
               ),
               
               actionButton(
                 "validar",
                 tags$span(icon("check-circle"), "Realizar Validação"),
                 class = "btn-validar"
               )
             ),
             
             mainPanel(
               tabsetPanel(
                 tabPanel(
                   tags$span(icon("chart-pie"), "Resumo"),
                   br(),
                   div(class = "resumo-box",
                       h4("Parâmetros da Validação"),
                       verbatimTextOutput("parametros"),
                       hr(),
                       h4("Resultados da Validação"),
                       verbatimTextOutput("resumo_validacao"),
                       uiOutput("download_button"),
                       DTOutput("resumo_tbl")
                   )
                 ),
                 tabPanel(
                   tags$span(icon("exclamation-triangle"), "Divergências"),
                   br(),
                   DTOutput("divergencias")
                 ),
                 tabPanel(
                   tags$span(icon("sync"), "Movimentação"),
                   br(),
                   DTOutput("movimentacao")
                 )
               )
             )
           )
  ),
  
  # Tab Manual
  tabPanel("Manual",
           fluidRow(
             column(
               8, offset = 2,
               div(
                 class = "manual-content",
                 style = "padding: 20px;",
                 h1("Manual do Usuário", align = "center"),
                 br(),
                 h2("1. Introdução"),
                 p("Esta aplicação foi desenvolvida para auxiliar na validação de bases de dados utilizando diferentes indexadores econômicos."),
                 
                 h2("2. Como Utilizar"),
                 h3("2.1 Seleção do Indexador"),
                 p("Escolha o indexador econômico desejado (IPCA, IGPM, INPC, SELIC ou TR)."),
                 
                 h3("2.2 Período de Análise"),
                 p("Defina o período de análise selecionando:"),
                 tags$ul(
                   tags$li("Data de Início: primeiro mês do período a ser analisado"),
                   tags$li("Data de Término: último mês do período a ser analisado")
                 ),
                 
                 h3("2.3 Carregamento das Bases"),
                 p("Faça o upload de duas bases de dados em formato Excel (.xlsx ou .xls):"),
                 tags$ul(
                   tags$li("Base 1: arquivo mais antigo"),
                   tags$li("Base 2: arquivo mais recente")
                 ),
                 
                 h3("2.4 Execução da Validação"),
                 p("Clique no botão 'Realizar Validação' para processar os dados."),
                 
                 h2("3. Resultados"),
                 p("Os resultados são apresentados em três abas:"),
                 tags$ul(
                   tags$li(strong("Resumo: "), "visão geral dos parâmetros e resultados da validação"),
                   tags$li(strong("Divergências: "), "lista detalhada de registros com diferenças"),
                   tags$li(strong("Movimentação: "), "análise das alterações entre as bases")
                 ),
                 
                 h2("4. Exportação"),
                 p("Use o botão 'Baixar Resultados' para exportar a análise completa em formato Excel.")
               )
             )
           )
  ),
  
  # Tab Sobre
  tabPanel("Sobre",
           fluidRow(
             column(
               8, offset = 2,
               div(
                 class = "sobre-content",
                 style = "padding: 20px; text-align: justify;",
                 h1("Sobre", align = "center"),
                 br(),
                 p("O GUISA - Sistema Atuarial é uma aplicação web desenvolvida como parte da pesquisa de conclusão do curso de Ciências Atuariais na 
     Universidade Federal da Paraíba (UFPB), por João Guilherme Pereira dos Santos, sob a orientação do Prof. Dr. Luiz Carlos Junior Santos.
     A aplicação foi projetada para oferecer uma ferramenta crítica e analítica das bases de dados gerenciadas por fundos de pensão no Brasil. 
     O objetivo principal do GUISA é agilizar e otimizar o processo de análise de dados, fornecendo, por meio de uma interface intuitiva e das 
     informações fornecidas pelo usuário, três abas principais:"),
                 tags$ul(
                   tags$li(tags$b("Dashboard Resumo"), " – Exibe uma análise prévia da massa de participantes, permitindo uma visão inicial dos dados."),
                   tags$li(tags$b("Tabela de Inconsistências"), " – Identifica e exibe divergências e inconsistências na base de dados entre períodos."),
                   tags$li(tags$b("Aba de Movimentação"), " – Mostra as entradas e saídas dos participantes ao longo do tempo.")
                 ),
                 p("Além disso, o sistema gera um relatório detalhado em formato .xlsx, permitindo uma visualização aprofundada das críticas e movimentações da base de dados.")
               )
             )
           )
  ),
  navbarMenu("Contato",
             tabPanel(HTML('<a href="https://mail.google.com/mail/?view=cm&fs=1&to=joaosantos34537445@gmail.com&su=Assunto&body=Corpo%20do%20email" target="_blank" style="font-size: 14px;">
              <img src="email.svg" width="14px" style="margin-right: 10px;"> E-mail
            </a>'),
                      tags$div(
                        class = "contact-item",
                        tags$a(
                          href = "mailto:seu.email@exemplo.com",
                          target = "_blank",
                          tags$img(src = "email.svg", width = "14px", style = "margin-right: 10px;"),
                          "E-mail"
                        )
                      )
             ),
             tabPanel(HTML('<a href="https://api.whatsapp.com/send?phone=5581999911222" target="_blank" style="font-size: 14px;">
              <img src="whatsapp.svg" width="14px" style="margin-right: 10px;"> WhatsApp
            </a>'),
                      tags$div(
                        class = "contact-item",
                        tags$a(
                          href = "https://api.whatsapp.com/send?phone=5581999911222",
                          target = "_blank",
                          tags$img(src = "whatsapp.svg", width = "14px", style = "margin-right: 10px;"),
                          "WhatsApp"
                        )
                      )
             ),
             tabPanel(HTML('<a href="https://github.com/jgpds" target="_blank" style="font-size: 14px;">
              <img src="github.svg" width="14px" style="margin-right: 10px;"> GitHub
            </a>'),
                      tags$div(
                        class = "contact-item",
                        tags$a(
                          href = "https://github.com/jgpds",
                          target = "_blank",
                          tags$img(src = "github.svg", width = "15px", style = "margin-right: 10px;"),
                          "GitHub"
                        )
                      )
             ),
             tabPanel(HTML('<a href="https://www.linkedin.com/in/joaoguilhermepsantos/" target="_blank" style="font-size: 14px;">
              <img src="linkedin.svg" width="17px" style="margin-right: 10px;"> LinkedIn
            </a>'),
                      tags$div(
                        class = "contact-item",
                        tags$a(
                          href = "https://www.linkedin.com/in/joaoguilhermepsantos/",
                          target = "_blank",
                          tags$img(src = "linkedin.svg", width = "17px", style = "margin-right: 10px;"),
                          "LinkedIn"
                        )
                      )
             )
  ),
  position = "fixed-top"
)
# Server
server <- function(input, output, session) {
  # Reactive values
  rv <- reactiveValues(
    validation_status = FALSE,
    current_params = NULL,
    validation_results = NULL,
    base1_data = NULL,
    base2_data = NULL
  )
  
  # Observer para ler Base 1
  observeEvent(input$base1, {
    req(input$base1)
    tryCatch({
      rv$base1_data <- read_excel(input$base1$datapath)
      showNotification("Base 1 carregada com sucesso!", type = "message")
    }, error = function(e) {
      showNotification(paste("Erro ao carregar Base 1:", e$message), type = "error")
    })
  })
  
  # Observer para ler Base 2
  observeEvent(input$base2, {
    req(input$base2)
    tryCatch({
      rv$base2_data <- read_excel(input$base2$datapath)
      showNotification("Base 2 carregada com sucesso!", type = "message")
    }, error = function(e) {
      showNotification(paste("Erro ao carregar Base 2:", e$message), type = "error")
    })
  })
  
  output$download_button <- renderUI({
    req(rv$validation_status)  # Só mostra o botão após a validação
    downloadButton(
      "downloadData",
      label = "Baixar Resultados",
      class = "btn-download"
    )
  })
  
  # Observe validation button click
  observeEvent(input$validar, {
    # Verificar se as bases foram carregadas
    if (is.null(rv$base1_data) || is.null(rv$base2_data)) {
      showNotification("Por favor, carregue ambas as bases antes de validar", type = "warning")
      return()
    }
    
    # Parse dates
    data_inicio <- as.Date(paste0(input$data_inicio, "-01"))
    data_fim <- as.Date(paste0(input$data_fim, "-01"))
    
    # Validate date range
    if (data_fim < data_inicio) {
      showNotification("Data de término deve ser posterior à data de início", type = "error")
      return()
    }
    
    # Get index data
    index_result <- get_accumulated_index(
      CONFIG$INDEXADORES[input$indexador],
      data_inicio,
      data_fim
    )
    
    if (!index_result$success) {
      showNotification(paste("Erro ao obter dados do indexador:", index_result$error), type = "error")
      return()
    }
    
    # Store parameters
    rv$current_params <- list(
      indexador = input$indexador,
      data_inicio = format(data_inicio, "%m/%Y"),
      data_fim = format(data_fim, "%m/%Y"),
      acumulado = index_result$value
    )
    
    rv$validation_status <- TRUE
  })
  
  # Download handler
  output$downloadData <- downloadHandler(
    filename = function() {
      paste0("resultados-", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      req(rv$current_params)
      
      wb <- createWorkbook()
      
      style_corpo_bordas <- createStyle(halign = "center", fontSize = 10, border = "TopBottomLeftRight", borderColour = "#A6A6A6", fontColour = "#262626")
      style_corpo_bordas_left <- createStyle(halign = "left", fontSize = 10, border = "TopBottomLeftRight", borderColour = "#A6A6A6", fontColour = "#262626")
      style_cabecalho2 <- createStyle(fontColour = "#262626", fgFill = "#D9D9D9", border = "TopBottomLeftRight", borderColour = "#A6A6A6",fontSize = 10, halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
      style_cabecalho3 <- createStyle(fontColour = "#262626", fgFill = "#ffde21", border = "TopBottomLeftRight",fontSize = 10, borderColour = "#A6A6A6", halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
      style_cabecalho4 <- createStyle(fontColour = "#f2eded", fgFill = "#474747", border = "TopBottomLeftRight", borderColour = "#A6A6A6",fontSize = 11, halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
      style_cabecalho5 <- createStyle(fontColour = "#f2eded", fgFill = "#474747", border = "TopBottomLeftRight", borderColour = "#f2eded",fontSize = 12, halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
      style_cabecalho <- createStyle(fontColour = "#262626", fgFill = "#d3d7de", border = "TopBottomLeftRight",fontSize = 10, halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
      
      
      # Aba de Parâmetros
      addWorksheet(wb, "Parâmetros")
      writeData(
        wb,
        "Parâmetros",
        data.frame(
          Parâmetro = c("Indexador", "Data Início", "Data Fim", "Acumulado (%)"),
          Valor = c(
            rv$current_params$indexador,
            rv$current_params$data_inicio,
            rv$current_params$data_fim,
            sprintf("%.2f%%", rv$current_params$acumulado * 100)
          )
        )
      )
      
      # Estilizar a aba de parâmetros
      headerStyle <- createStyle(
        textDecoration = "bold",
        border = "Bottom",
        borderStyle = "medium"
      )
      addStyle(wb, "Parâmetros", headerStyle, rows = 1, cols = 1:2)
      setColWidths(wb, "Parâmetros", cols = 1:2, widths = "auto")
      
      # Aba Base 1
      if (!is.null(rv$base1_data)) {
        addWorksheet(wb, "Base 1")
        writeData(wb, "Base 1", rv$base1_data)
        setColWidths(wb, "Base 1", cols = 1:ncol(rv$base1_data), widths = "auto")
        
        addStyle(wb, sheet = 2, style_cabecalho2, rows = 1, cols = 1:ncol(rv$base1_data))
        addStyle(wb, sheet = 2, style_corpo_bordas, rows = 2:(1+nrow(rv$base1_data)), cols = 1:ncol(rv$base1_data), gridExpand = TRUE)
        
        
      }
      
      # Aba Base 2
      if (!is.null(rv$base2_data)) {
        addWorksheet(wb, "Base 2")
        writeData(wb, "Base 2", rv$base2_data)
        setColWidths(wb, "Base 2", cols = 1:ncol(rv$base2_data), widths = "auto")
        
        addStyle(wb, sheet = 3, style_cabecalho2, rows = 1, cols = 1:ncol(rv$base2_data))
        addStyle(wb, sheet = 3, style_corpo_bordas, rows = 2:(1+nrow(rv$base2_data)), cols = 1:ncol(rv$base2_data), gridExpand = TRUE)
      }
      
      # Salvar o workbook
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
}

# Run the application
shinyApp(ui = ui, server = server)