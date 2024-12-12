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
  dplyr,
  plotly,
  ggplot2,
  lookup
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
      body {
        overflow-y: auto; /* Remove a barra de rolagem vertical */
      }
      .navbar-page {
        max-width: 800px; /* Define a largura máxima */
        max-height: 600px; /* Define a altura máxima */
        overflow-y: auto; /* Permite rolagem interna nos painéis se necessário */
      }
      .tab-content {
        max-height: 500px; /* Altura máxima do conteúdo de cada aba */
      }
    ")),
    tags$link(rel = "shortcut icon", href = "favicon.ico")
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
                 choices = c(1, 2, 3, 4, 5, 10),
                 multiple = TRUE,
                 selected = 5
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
               ),
               
               # Adicionando o botão de download aqui
               uiOutput("download_button")
             ),
             
             mainPanel(
               tabsetPanel(
                 tabPanel(
                   tags$span(icon("chart-pie"), "Resumo"),
                   br(),
                   fluidRow(
                     column(3,
                            div(class = "info-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4(icon("users"), "Total Participantes", style = "font-size: 14px; margin: 0;"),
                                div(style = "font-size: 24px; font-weight: bold; color: #2c3e50;",
                                    textOutput("total_participantes")
                                )
                            )
                     ),
                     column(3,
                            div(class = "info-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4(icon("venus-mars"), "Distribuição por Sexo", style = "font-size: 14px; margin: 0;"),
                                div(style = "font-size: 24px; font-weight: bold; color: #2c3e50;",
                                    textOutput("dist_sexo")
                                )
                            )
                     ),
                     column(3,
                            div(class = "info-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4(icon("calendar"), "Idade Média", style = "font-size: 14px; margin: 0;"),
                                div(style = "font-size: 24px; font-weight: bold; color: #2c3e50;",
                                    textOutput("idade_media")
                                )
                            )
                     ),
                     column(3,
                            div(class = "info-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4(icon("dollar-sign"), "Benefício Médio", style = "font-size: 14px; margin: 0;"),
                                div(style = "font-size: 24px; font-weight: bold; color: #2c3e50;",
                                    textOutput("beneficio_medio")
                                )
                            )
                     )
                   ),
                   
                   # Gráficos (mantidos apenas os dois primeiros)
                   fluidRow(
                     column(6,
                            div(class = "chart-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4("Pirâmide Etária", style = "font-size: 16px; margin-bottom: 15px;"),
                                plotlyOutput("piramide_etaria", height = "500px")
                            )
                     ),
                     column(6,
                            div(class = "chart-box",
                                style = "padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px;",
                                h4("Distribuição Benefício", style = "font-size: 16px; margin-bottom: 15px;"),
                                plotlyOutput("dist_beneficio", height = "500px")
                            )
                     )
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
                 p("Use o botão abaixo para baixar o modelo de planilhas:"),
                 downloadButton("download_planilhas", "Baixar Modelos de Planilhas")
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
             HTML('<li><a href="https://mail.google.com/mail/?view=cm&fs=1&to=joaosantos34537445@gmail.com&su=Assunto&body=Corpo%20do%20email" target="_blank" style="font-size: 14px;">
                  <img src="email.svg" width="14px" style="margin-right: 10px;"> E-mail</a></li>'),
             HTML('<li><a href="https://api.whatsapp.com/send?phone=5581999911222" target="_blank" style="font-size: 14px;">
                  <img src="whatsapp.svg" width="14px" style="margin-right: 10px;"> WhatsApp</a></li>'),
             HTML('<li><a href="https://github.com/jgpds" target="_blank" style="font-size: 14px;">
                  <img src="github.svg" width="14px" style="margin-right: 10px;"> GitHub</a></li>'),
             HTML('<li><a href="https://www.linkedin.com/in/joaoguilhermepsantos/" target="_blank" style="font-size: 14px;">
                  <img src="linkedin.svg" width="17px" style="margin-right: 10px;"> LinkedIn</a></li>')
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
  
  # Modificação no server para o botão de download
  output$download_button <- renderUI({
    if (rv$validation_status) {
      downloadButton(
        "downloadData",
        label = "Baixar Relatório ",
        class = "btn-download btn-block", # Adicionado btn-block para largura total
        style = "margin-top: 15px; width: 100%;" # Aumentado margin-top e definido width 100%
      )
    }
  })
  
  # output$piramide_etaria <- renderPlotly({
  #   req(rv$base2_data)
  #   dados <- rv$base2_data
  #   datas_nascimento <- as.Date(dados$`Data de Nascimento`, origin = "1899-12-30")
  #   data_atual <- as.Date(input$data_fim, origin = "1899-12-30") # mesma base de data
  #   idades <- as.numeric(difftime(data_atual, datas_nascimento, units = "weeks")) %/% 52
  #   
  #   pyramid_data <- data.frame(
  #     Edad = abs(idades),
  #     Sexo = ifelse(rv$base2_data$Sexo == "Masculino", "Hombres", "Mujeres"),
  #     Cantidad = 1
  #   )
  #   
  #   ggplot(pyramid_data, aes(x = Edad, fill = Sexo)) +
  #     geom_bar(aes(y = Cantidad), stat = "identity", position = "stack") +
  #     scale_x_reverse() +
  #     labs(x = "Edad", y = "Población", title = "Pirámide Poblacional") +
  #     theme_classic() +
  #     theme(
  #       axis.text.y = element_text(size = 10),
  #       axis.text.x = element_text(size = 8),
  #       plot.title = element_text(size = 12, face = "bold")
  #     )
  # })
  
  output$piramide_etaria <- renderPlotly({
    req(rv$base2_data)
    dados <- rv$base2_data
    datas_nascimento <- as.Date(dados$`Data de Nascimento`, origin = "1899-12-30")
    data_atual <- as.Date(input$data_fim, origin = "1899-12-30") # mesma base de data
    idades <- as.numeric(difftime(data_atual, datas_nascimento, units = "weeks")) %/% 52
    sexo <- dados$Sexo
    
    # Criar um dataframe com as idades e sexo
    dados <- data.frame(idade = idades, sexo = sexo)
    
    # Criar faixas etárias (0-4, 5-9, ..., 95-99) e contar por faixa etária e sexo
    dados <- dados %>%
      mutate(faixa_etaria = cut(idade, breaks = seq(0, 100, by = 3), right = FALSE)) %>%
      group_by(faixa_etaria, sexo) %>%
      summarise(count = n(), .groups = "drop") %>%
      mutate(count = ifelse(sexo == "M", -count, count))
    
    # Criar o gráfico de pirâmide etária
    plot_ly(dados, x = ~count, y = ~faixa_etaria, color = ~sexo, colors = c("#66C3FF", "pink"),
            type = 'bar', orientation = 'h') %>%
      layout(barmode = 'relative',
             xaxis = list(title = 'População', tickvals = seq(-max(abs(dados$count)), max(abs(dados$count)), by = 5),
                          ticktext = abs(seq(-max(abs(dados$count)), max(abs(dados$count)), by = 5))),
             yaxis = list(title = 'Faixa Etária'),
             title = "Pirâmide Etária")
  })
  
  
  output$dist_beneficio <- renderPlotly({
    req(rv$base2_data)
    
    tryCatch({
      # Extrai e limpa os dados de benefícios
      dados <- rv$base2_data
      beneficios <- as.numeric(gsub("[^0-9.]", "", dados[[6]]))
      
      # Remove NAs e valores negativos ou zero
      beneficios <- beneficios[!is.na(beneficios) & beneficios > 0]
      
      # Cria o boxplot básico com ggplot2
      p <- ggplot(data.frame(Beneficio = beneficios), aes(y = Beneficio)) +
        geom_boxplot(outlier.colour = "black", fill = "#e6e7e8") +
        labs(title = "", y = "Valor do Benefício (R$)") +
        theme_minimal()
      
      # Converte para um gráfico interativo com ggplotly
      ggplotly(p)
      
    }, error = function(e) {
      plot_ly() %>%
        add_annotations(
          text = "Erro ao gerar o gráfico. Verifique os dados.",
          showarrow = FALSE
        )
    })
  })
  
  # Calcular o total de participantes
  output$total_participantes <- renderText({
    total <- nrow(rv$base2_data)
    return(total)
  })
  
  output$idade_media <- renderText({
    # Datas fornecidas
    datas_nascimento <- as.Date(rv$base2_data[[2]], origin = "1899-12-30") # para o sistema de data do Excel (base 1900)
    data_atual <- as.Date(input$data_fim, origin = "1899-12-30") # mesma base de data
    
    # Calcular idade
    idade_media_aux <- mean(as.numeric(difftime(data_atual, datas_nascimento, units = "weeks")) %/% 52)
    return(idade_media_aux)
  })
  
  output$beneficio_medio <- renderText({
    req(rv$base2_data)  # Garante que os dados existem
    
    # Pega o nome da coluna de benefício (ajuste conforme sua base)
    coluna_beneficio <- names(rv$base2_data)[6]
    
    # Converte para numérico, removendo caracteres não numéricos
    beneficios <- as.numeric(gsub("[^0-9.]", "", rv$base2_data[[6]]))
    
    # Calcula a média, removendo NAs
    ben_medio <- mean(beneficios, na.rm = TRUE)
    
    # Formata o resultado
    return(format(round(ben_medio, 2), big.mark = ".", decimal.mark = ","))
  })
  
  output$dist_sexo <- renderText({
    req(rv$base2_data)  # Garante que os dados existem
    
    num_homens <- sum(ifelse(rv$base2_data[[3]] == "M", 1, 0))
    num_mulheres <- sum(ifelse(rv$base2_data[[3]] == "F", 1, 0))
    
    # Formata o resultado
    return(paste0("Masculino: ", num_homens, "\nFeminimo: ", num_mulheres))
  })
  
  # Link download sobre
  output$download_planilhas <- downloadHandler(
    filename = function() {
      paste("Modelos-Planilhas-", Sys.Date(), ".zip", sep = "")
    },
    content = function(file) {
      # Diretório dos arquivos existentes
      file1 <- "datasets/Base Anterior.xlsx"
      file2 <- "datasets/Base Atual.xlsx"
      
      # Compactando os arquivos em um único ZIP
      zip::zipr(file, files = c(file1, file2))
    }
  )
  
  
  
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
      req(rv$base2_data)
      req(rv$base1_data)
      
      dados <- rv$base2_data
      datas_nascimento <- as.Date(dados$`Data de Nascimento`, origin = "1899-12-30")
      data_atual <- as.Date(input$data_fim, origin = "1899-12-30") # mesma base de data
      idades <- as.numeric(difftime(data_atual, datas_nascimento, units = "weeks")) %/% 52
      sexo <- dados$Sexo
      
      # Salário Mínimo
      sal_minimo <- rbcb::get_series(1619
                                     , start_date = input$data_inicio, end_date = input$data_fim)
      sal_minimo <- tail(sal_minimo$`1619`, n = 1)
      
      base_anterior <- rv$base1_data      
      base_atual <- rv$base2_data
      
      # Interseção
      matrículas_comuns  <- intersect(base_atual$Matrícula, base_anterior$Matrícula)
      base_atual_x_anterior <- data.frame(Matrículas = matrículas_comuns)
      
      # Bases Auxiliares
      divergencias_atual <- tibble(                                             
        Matrícula = base_atual$Matrícula,
        `Matrículas Repetidas` = ifelse(duplicated(base_atual$Matrícula) | duplicated(base_atual$Matrícula, fromLast = TRUE), 1, 0),
        `Idade maior que 90 anos` = ifelse(idades > 90, 1, 0),
        `Benefício menor que o salário mínimo` = ifelse(base_atual$`Benefício Complementar` < sal_minimo, 1, 0),
        `Data de Início de Benefício menor que a Data de Nascimento` = rep(NA,nrow(base_atual)))
      
      divergencias_atual_anterior <- tibble(                                             
        Matrícula = base_atual_x_anterior$`Matrículas`,
        `Data de Nascimento divergente` = rep(NA,nrow(base_atual_x_anterior)),
        `Sexo divergente` = rep(NA,nrow(base_atual_x_anterior)),
        `Tipo de Benefício divergente` = rep(NA,nrow(base_atual_x_anterior)),
        `Benefício Complementar com variações significativas` = rep(NA,nrow(base_atual_x_anterior)))
      
      # Críticas
      criticas <- data.frame(
        Critica = c(
          "Matrículas Repetidas", 
          "Idade maior que 90 anos", 
          "Benefício menor que o salário mínimo", 
          "Data de Início de Benefício menor que a Data de Nascimento", 
          "Data de Nascimento divergente", 
          "Sexo divergente", 
          "Tipo de Benefício divergente", 
          "Benefício Complementar com variações significativas"
        ),
        Frequencia = NA,      # Initialize with NA as in the original table
        Percentual_Sobre_a_Base = NA  # Initialize with NA as in the original table
      )
      colnames(criticas) <- c("Críticas", "Frequência", "Percentual Sobre a Base")
      
      # 1. Matrículas Repetidas
      
      criticas[1, 2] <- sum(ifelse(duplicated(rv$base2_data$Matrícula) | duplicated(rv$base2_data$Matrícula, fromLast = TRUE), 1, 0)) # Ocorrëncias
      criticas[1, 3] <- paste0(round(criticas[1, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      # 2. Idade maior que 90 anos
      criticas[2, 2] <- sum(ifelse(idades > 90, 1, 0)) # Ocorrëncias
      criticas[2, 3] <- paste0(round(criticas[2, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      # 3. Benefício menor que o salário mínimo
      criticas[3, 2] <- sum(ifelse(base_atual$`Benefício Complementar` < sal_minimo, 1, 0)) # Ocorrëncias
      criticas[3, 3] <- paste0(round(criticas[3, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      # Data de Início de Benefício menor que a Data de Nascimento
      
      # Data de Nascimento divergente
      atual <- vlookup(base_atual_x_anterior$`Matrícula`, base_atual, "Matrícula", "Data de Nascimento")
      anterior <- vlookup(base_atual_x_anterior$`Matrícula`, base_anterior, "Matrícula", "Data de Nascimento")
      divergencias_atual_anterior$`Data de Nascimento divergente` <- ifelse(atual != anterior, 1, 0)
      criticas[5, 2] <- sum(divergencias_atual_anterior$`Data de Nascimento divergente`) # Ocorrëncias
      criticas[5, 3] <- paste0(round(criticas[5, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      # Sexo divergente
      atual <- vlookup(base_atual_x_anterior$`Matrícula`, base_atual, "Matrícula", "Sexo")
      anterior <- vlookup(base_atual_x_anterior$`Matrícula`, base_anterior, "Matrícula", "Sexo")
      divergencias_atual_anterior$`Sexo divergente` <- ifelse(atual != anterior, 1, 0)
      criticas[6, 2] <- sum(ifelse(atual != anterior, 1, 0)) # Ocorrëncias
      criticas[6, 3] <- paste0(round(criticas[6, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      # Tipo de Benefício divergente
      atual <- vlookup(base_atual_x_anterior$`Matrícula`, base_atual, "Matrícula", "Tipo do Benefício")
      anterior <- vlookup(base_atual_x_anterior$`Matrícula`, base_anterior, "Matrícula", "Tipo do Benefício")
      divergencias_atual_anterior$`Tipo de Benefício divergente` <- ifelse(atual != anterior, 1, 0)
      criticas[7, 2] <- sum(ifelse(atual != anterior, 1, 0)) # Ocorrëncias
      criticas[7, 3] <- paste0(round(criticas[7, 2]/nrow(base_atual) * 100, 2), "%") # Porcentagem sobre a base
      
      
      # Anexo
      
      # Código 1
      df1_anexoI_titulo <- "Código 1\nMatrículas Repetidas"
      df1_anexoI <- data.frame(Matrícula = (divergencias_atual %>% filter(`Matrículas Repetidas` == 1))$Matrícula)
      
      # Código 2
      df2_anexoI_titulo <- "Código 2\nIdade maior que 90 anos"
      df2_anexoI <- data.frame(Matrícula = (divergencias_atual %>% filter(`Idade maior que 90 anos` == 1))$Matrícula)
      df2_anexoI$`Data de Nascimento` <- vlookup(df2_anexoI$`Matrícula`, base_atual, "Matrícula", "Data de Nascimento")
      df2_anexoI$`Idade` <- as.numeric(difftime(data_atual, df2_anexoI$`Data de Nascimento`, units = "weeks")) %/% 52
      
      # Código 3
      df3_anexoI_titulo <- "Código 3\nBenefício menor que o salário mínimo"
      df3_anexoI <- data.frame(Matrícula = (divergencias_atual %>% filter(`Benefício menor que o salário mínimo` == 1))$Matrícula)
      df3_anexoI$`Benefício Complementar` <- vlookup(df3_anexoI$`Matrícula`, base_atual, "Matrícula", "Benefício Complementar")
      df3_anexoI$`Salário mínimo` <- sal_minimo
      # Código 4
      
      # Código 5
      df5_anexoI_titulo <- "Código 5\nData de Nascimento divergente"
      df5_anexoI <- data.frame(Matrícula = (divergencias_atual_anterior %>% filter(`Data de Nascimento divergente` == 1))$Matrícula)
      df5_anexoI$`Data de Nascimento Anterior` <- vlookup(df5_anexoI$`Matrícula`, base_atual, "Matrícula", "Data de Nascimento")
      df5_anexoI$`Data de Nascimento Atual` <- vlookup(df5_anexoI$`Matrícula`, base_anterior, "Matrícula", "Data de Nascimento")
      
      # Código 6
      df6_anexoI_titulo <- "Código 6\nSexo divergente"
      df6_anexoI <- data.frame(Matrícula = (divergencias_atual_anterior %>% filter(`Sexo divergente` == 1))$Matrícula)
      
      # Código 6
      df7_anexoI_titulo <- "Código 7\nTipo de Benefício divergente"
      df7_anexoI <- data.frame(Matrícula = (divergencias_atual_anterior %>% filter(`Tipo de Benefício divergente` == 1))$Matrícula)
      
      # Auxílio para loop ao preencher excel
      dfs_anexoI_títulos <- c(df1_anexoI_titulo, df2_anexoI_titulo, df3_anexoI_titulo, df5_anexoI_titulo, df6_anexoI_titulo, df7_anexoI_titulo)
      dfs_anexoI <- list(df1_anexoI, df2_anexoI, df3_anexoI, df5_anexoI, df6_anexoI, df7_anexoI)
      
      
      
      # Criando pasta excel
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
        writeData(wb, "Base 1", rv$base1_data, startCol = 2, startRow = 2)
        setColWidths(wb, "Base 1", cols = 1, widths = 2.5)
        setColWidths(wb, "Base 1", cols = 2:ncol(rv$base1_data), widths = "auto")
        
        addStyle(wb, sheet = 2, style_cabecalho2, rows = 2, cols = 2:(1+ncol(rv$base1_data)))
        addStyle(wb, sheet = 2, style_corpo_bordas, rows = 3:(2+nrow(rv$base1_data)), cols = 2:(1+ncol(rv$base1_data)), gridExpand = TRUE)
        
        
      }
      
      # Aba Base 2
      if (!is.null(rv$base2_data)) {
        addWorksheet(wb, "Base 2")
        writeData(wb, "Base 2", rv$base2_data, startCol = 2, startRow = 2)
        setColWidths(wb, "Base 2", cols = 1, widths = 2.5)
        setColWidths(wb, "Base 2", cols = 2:ncol(rv$base2_data), widths = "auto")
        
        addStyle(wb, sheet = "Base 2", style_cabecalho2, rows = 2, cols = 2:(1+ncol(rv$base2_data)))
        addStyle(wb, sheet = "Base 2", style_corpo_bordas, rows = 3:(2+nrow(rv$base2_data)), cols = 2:(1+ncol(rv$base2_data)), gridExpand = TRUE)
      }
      
      # Aba de Parâmetros
      addWorksheet(wb, "Movimentação")
      
      # Aba Divergências
      addWorksheet(wb, "Divergências")
      writeData(wb, "Divergências", "Resumo - Crítica da Base de Dados", startCol = 2, startRow = 2)
      writeData(wb, "Divergências", criticas, startCol = 2, startRow = 3)
      
      mergeCells(wb, sheet = "Divergências", cols = 2:(1+ncol(criticas)), rows = 2)
      
      setColWidths(wb, sheet = "Divergências", cols = 2, widths = 57) 
      setColWidths(wb, sheet = "Divergências", cols = 3, widths = 11) 
      setColWidths(wb, sheet = "Divergências", cols = 4, widths = 21) 
      
      addStyle(wb, sheet = "Divergências", style_cabecalho4, rows = 2, cols = 2:(1+ncol(criticas)), gridExpand = TRUE)
      addStyle(wb, sheet = "Divergências", style_cabecalho2, rows = 3, cols = 2:(1+ncol(criticas)), gridExpand = TRUE)
      addStyle(wb, sheet = "Divergências", style_corpo_bordas_left, rows = 4:(3+nrow(criticas)), cols = 2, gridExpand = TRUE)
      addStyle(wb, sheet = "Divergências", style_corpo_bordas, rows = 4:(3+nrow(criticas)), cols = 3:(1+ncol(criticas)), gridExpand = TRUE)
      # Aba Anexo
      addWorksheet(wb, "Anexo")
      
      starting_col <- 2
      for (i in 1:length(dfs_anexoI_títulos)){
        if (nrow(dfs_anexoI[[i]]) > 0 ){
          writeData(wb, sheet = "Anexo", x = dfs_anexoI_títulos[i], startCol = starting_col, startRow = 2)
          writeData(wb, sheet = "Anexo", x = dfs_anexoI[[i]], startCol = starting_col, startRow = 3)
          
          mergeCells(wb, sheet = "Anexo", cols = starting_col:(starting_col+length(dfs_anexoI[[i]])-1), rows = 2) # Unindo colunas do título
          
          setColWidths(wb, sheet = "Anexo", cols = starting_col, widths = 14) # Largura primeira coluna (Matricula)
          setColWidths(wb, sheet = "Anexo", cols = (starting_col+1):(starting_col+length(dfs_anexoI[[i]])-1), widths = 22) # Largura resto dos subtítulos
          
          addStyle(wb, sheet = "Anexo", style_cabecalho4, rows = 2, cols = starting_col:(starting_col+length(dfs_anexoI[[i]])-1))
          addStyle(wb, sheet = "Anexo", style_cabecalho2, rows = 3, cols = starting_col:(starting_col+length(dfs_anexoI[[i]])-1))
          addStyle(wb, sheet = "Anexo", style_corpo_bordas, rows = 4:(3+nrow(dfs_anexoI[[i]])), cols = starting_col:(starting_col+length(dfs_anexoI[[i]])-1), gridExpand = TRUE)
          
          setColWidths(wb, sheet = "Anexo", cols = starting_col + length(dfs_anexoI[[i]]) + 1, widths = 1)
          starting_col <- starting_col + length(dfs_anexoI[[i]]) + 2
        }
      }
      
      setColWidths(wb, sheet = "Anexo", cols = 1, widths = 1)
      setRowHeights(wb, sheet = "Anexo", rows = 1, heights = 9)
      setRowHeights(wb, sheet = "Anexo", rows = 2, heights = 44)
      
      # Salvar o workbook
      saveWorkbook(wb, file, overwrite = TRUE)
      rv$criticas <- criticas
    }
  )
  
  output$divergencias <- renderDT({
    req(rv$criticas)
    
    datatable(
      rv$criticas,
      options = list(
        pageLength = 8,
        dom = 'rtip',
        language = list(
          info = "Mostrando _START_ até _END_ de _TOTAL_ registros"
        )
      ),
      class = 'cell-border stripe hover',
      rownames = FALSE
    )
  })
  
}

# Run the application
shinyApp(ui = ui, server = server)