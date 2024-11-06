server <- function(input, output, session) {
  # Create reactive value to track validation status
  validation_status <- reactiveVal(FALSE)
  
  # Função reativa para capturar os parâmetros ao clicar no botão
  parametros_reativos <- eventReactive(input$validar, {
    # Extrair o ano e o mês a partir do airMonthpickerInput
    data_inicio <- as.Date(paste(substr(input$data_inicio, 1, 4), 
                                 substr(input$data_inicio, 6, 7), "01", sep = "-"))
    data_fim <- as.Date(paste(substr(input$data_fim, 1, 4), 
                              substr(input$data_fim, 6, 7), "01", sep = "-"))
    
    parametros <- list(
      "Data de Início" = format(data_inicio, "%m/%Y"),
      "Data de Término" = format(data_fim, "%m/%Y"),
      "Indexador Econômico" = input$indexador
    )
    
    # Exportar os parâmetros para um arquivo Excel
    write.xlsx(as.data.frame(parametros), "resultados.xlsx", rowNames = FALSE)
    
    # Set validation status to TRUE to show download button
    validation_status(TRUE)
    
    return(parametros)
  })
  
  # Render download button conditionally
  output$download_button <- renderUI({
    if (validation_status()) {
      downloadButton("downloadData", 
                     label = "Relatório Detalhado",  # Removido o tags$span e icon()
                     class = "btn btn-success"
      )
    }
  })
  
  # Handle file download
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("resultados-", Sys.Date(), ".xlsx", sep="")
    },
    content = function(file) {
      file.copy("resultados.xlsx", file)
    }
  )
  
  
  # Exibir os parâmetros capturados no verbatimTextOutput
  output$parametros <- renderText({
    paste(
      "Parâmetros salvos no arquivo 'resultados.xlsx':",
      "\n",
      parametros_reativos()$`Data de Início`,
      parametros_reativos()$`Data de Término`,
      parametros_reativos()$`Indexador Econômico`
    )
  })
  
  # Outras funções do servidor permanecem iguais...
}

shinyApp(ui = ui, server = server)