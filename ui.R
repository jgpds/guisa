ui <- fluidPage(
  theme = bs_theme(
    bg = "#ffffff",
    fg = "#333333",
    primary = "#575152",
    base_font = "Montserrat",
    heading_font = "Montserrat",
    font_scale = 1.1
  ),
  
  tags$head(
    tags$style(HTML("
      .titulo-app {
        padding: 10px;
        text-align: left;
        background-color: #ffffff;
        color: white;
        margin-bottom: 20px;
        border-radius: 5px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      
      .titulo-principal {
        font-size: 36px;
        font-weight: 900;
        letter-spacing: 2px;
        margin: 0;
        line-height: 1;
      }
      
      .subtitulo {
        font-size: 6px;
        letter-spacing: 6px;
        margin-top: 5px;
        font-weight: 350;
      }
      
      /* Resto dos estilos permanece igual... */
      .file-input {
        margin-bottom: 20px;
      }
      
      .input-group {
        margin-bottom: 20px;
      }
      
      .btn-validar {
        width: 100%;
        margin-top: 20px;
        padding: 10px;
        font-weight: bold;
        text-transform: uppercase;
        background-color: #575152;
        border-color: #404040;
        color: #ffffff;
      }
      
      .btn-validar:hover {
        background-color: #404040;
        border-color: #303030;
      }
      
      .well {
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-radius: 8px;
        background-color: #f8f9fa;
      }
      
      .tab-content {
        padding: 20px;
        background-color: #ffffff;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      
      .nav-tabs > li > a {
        color: #575152;
      }
      
      .dataTables_wrapper {
        padding: 20px;
        background-color: #ffffff;
        border-radius: 8px;
      }
      
      .menu-buttons {
        display: flex;
      }
      
      .menu-buttons a {
        color: #333;
        text-decoration: none;
        padding: 10px 40px;
        position: relative;
        transition: color 0.3s ease;
        font-size: 11px;
        letter-spacing: 6px;
      }
      
      .menu-buttons a:before {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 0;
        height: 2px;
        background-color: #575152;
        transition: width 0.3s ease;
      }
      
      .menu-buttons a:hover {
        color: #575152;
      }
      
      .menu-buttons a:hover:before {
        width: 100%;
      }
      
      .form-group {
        margin-bottom: 15px;
      }
      .dropdown {
        position: relative;
        display: inline-block;
      }
      
      .dropdown-btn {
        color: #333;
        text-decoration: none;
        padding: 10px 20px;
        position: relative;
        transition: color 0.3s ease;
        font-size: 14px;
        cursor: pointer;
      }
      
      .dropdown-content {
        display: none;
        position: absolute;
        background-color: #f1f1f1;
        min-width: 175px;
        box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
        z-index: 1;
      }
      
      .dropdown-content a {
        color: #333;
        padding: 12px 16px;
        text-decoration: none;
        display: block;
        font-size: 11px;
        letter-spacing: 0px;
        text-align: center;
      }
      
      .dropdown-content a:hover {
        background-color: #ddd;
      }
      
      .dropdown:hover .dropdown-content {
        display: block;
      }
      .air-datepicker-cell.-month- {
       color: #333333 !important; /* Define a cor padrão dos meses */
      }
      .air-datepicker-cell.-selected- {
        background-color: #575152 !important;
        color: #ffffff !important;
      }
      .contato img {
        width: 20px;
        margin-right: 5px;
        vertical-align: middle;
      }
      .divider {
        margin: 0 15px;  /* espaçamento horizontal */
        color: #333;     /* cor da divisória */
        opacity: 0.5;    /* transparência opcional */
      }
        .btn-success {
        margin-top: 15px;
        margin-bottom: 15px;
        background-color: #28a745;
        border-color: #28a745;
        }
        .btn-success:hover {
        background-color: #218838;
        border-color: #1e7e34;
      }
    "))
  ),
  
  div(class = "titulo-app",
      tags$img(src = "Screenshot_2.png", height = "80px"),
      div(class = "menu-buttons",
          a(href="#", "MANUAL"),
          span(class = "divider", "|"),  # divisória
          a(href="#", "SOBRE"),
          span(class = "divider", "|"),  # divisória
          div(class = "dropdown",
              a(href="#", class = "dropdown-btn", "CONTATO"),
              div(class = "dropdown-content",
                a(href="#", tags$img(src = "email.svg", target = "_blank", width = "14px", class = "dropdown-icon"), "E-mail"),
                a(href="https://api.whatsapp.com/send?phone=5581999911222", target = "_blank", tags$img(src = "whatsapp.svg", width = "14px", class = "dropdown-icon"), "WhatsApp"),
                a(href="https://github.com/jgpds", target = "_blank", tags$img(src = "github.svg", width = "15px", class = "dropdown-icon"), "GitHub"),
                a(href="#", tags$img(src = "linkedin.svg", target = "_blank", width = "17px", class = "dropdown-icon"), "Linkedin")
              )
          )
      )
  ),
  
  # Resto do código UI permanece igual...
  sidebarLayout(
    sidebarPanel(
      selectInput("indexador", 
                  tags$span(icon("chart-line"), "Indexador Econômico"),
                  choices = c("IPCA", "IGPM", "INPC", "SELIC", "TR"),
                  selected = "IPCA"),
      
      airMonthpickerInput("data_inicio",
                          tags$span(icon("calendar"), "Data de Início"),
                          minDate = "2000-01",
                          maxDate = "2030-12",
                          view = "months",
                          dateFormat = "MM/yyyy"),
      
      airMonthpickerInput("data_fim",
                          tags$span(icon("calendar-check"), "Data de Término"),
                          minDate = "2000-01",
                          maxDate = "2030-12",
                          view = "months",
                          dateFormat = "MM/yyyy"),
      
      selectizeInput("margem-tolerância", 
                     tags$span(icon("sliders-h"), "Margem de Tolerância (%):"), 
                     choices = NULL, 
                     multiple = TRUE),
      
      hr(),
      
      div(class = "file-input",
          fileInput("base1", 
                    tags$span(icon("file-excel"), "Base 1 (Mais antiga)"),
                    accept = c(".xlsx", ".xls"))
      ),
      
      div(class = "file-input",
          fileInput("base2", 
                    tags$span(icon("file-excel"), "Base 2 (Mais recente)"),
                    accept = c(".xlsx", ".xls"))
      ),
      
      actionButton("validar", 
                   tags$span(icon("check-circle"), "Realizar Validação"),
                   class = "btn-validar")
    ),
    
    mainPanel(
      tabsetPanel(
        tabPanel(tags$span(icon("chart-pie"), "Resumo"), 
                 br(),
                 div(class = "resumo-box",
                     h4("Parâmetros da Validação"),
                     verbatimTextOutput("parametros"),
                     hr(),
                     h4("Resultados da Validação"),
                     verbatimTextOutput("resumo_validacao"),
                     # Add conditional download button
                     uiOutput("download_button"),
                     DTOutput("resumo_tbl"))
        ),
        tabPanel(tags$span(icon("exclamation-triangle"), "Divergências"), 
                 br(),
                 DTOutput("divergencias")),
        tabPanel(tags$span(icon("sync"), "Movimentação"), 
                 br(),
                 DTOutput("movimentacao"))
      )
    )
  )
)