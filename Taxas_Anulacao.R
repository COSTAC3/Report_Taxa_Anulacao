# Instalar e carregar os pacotes necessários

library(readxl)
library(writexl)

# Diretório e nome do ficheiro original
input_file <- "C:/Users/COSTAC3/Desktop/Teste_Anulações/Report_Taxa_anulacao_Cognos.xlsx"

# Diretório e nome do ficheiro modificado
output_file <- "C:/Users/COSTAC3/Desktop/Teste_Anulações/REPORT_TAXAS_ANULACOES.xlsx"

# Verificar se o arquivo original existe
if (!file.exists(input_file)) {
  stop(paste("O arquivo especificado não existe. Verifique o caminho:", input_file))
}

# Verificar se o arquivo modificado já existe e está acessível
if (file.exists(output_file)) {
  # Tentar abrir o arquivo para escrita
  is_accessible <- tryCatch({
    fileConn <- file(output_file, open = "w") # Tenta abrir o arquivo para escrita
    close(fileConn) # Fecha o arquivo
    TRUE # Se conseguir abrir, retorna TRUE
  }, error = function(e) {
    FALSE # Se houver erro, retorna FALSE
  })
  
  if (!is_accessible) {
    cat("ATENÇÃO: O arquivo", output_file, "já existe e está aberto.\n")
    cat("Certifique-se de que o arquivo não está aberto e tente novamente.\n")
    stop("O código foi interrompido para evitar problemas ao salvar o arquivo.")
  }
}

# Listar todas as folhas disponíveis no arquivo
sheet_names <- excel_sheets(input_file)
cat("Folhas disponíveis no ficheiro:\n", sheet_names)


# Ler os dados da folha "Page1_1"
data <- tryCatch({
  read_excel(input_file, sheet = "Page1_1", col_names = FALSE)
}, error = function(e) {
  stop("Erro ao ler os dados da folha 'Page1_1'. Detalhes do erro: ", e$message)
})

# Verificar se os dados foram carregados corretamente
if (is.null(data)) {
  stop("Não foi possível carregar os dados da folha 'Page1_1'.")
}

# Função para preencher células vazias com o valor da célula acima
fill_empty_cells <- function(df) {
  for (col in seq_len(ncol(df))) { # Para cada coluna
    for (row in seq_len(nrow(df))) { # Para cada linha
      if ((row > 1) && (is.na(df[row, col]) || df[row, col] == "")) { # Verificar se a célula está vazia
        df[row, col] <- df[row - 1, col] # Preencher com o valor da célula acima
      }
    }
  }
  return(df)
}

# Aplicar a função para preencher células vazias
data <- fill_empty_cells(data)

# Remover as 6 primeiras linhas e as 2 últimas
if (nrow(data) > 8) {
  data <- data[-c(1:6, (nrow(data)-1):nrow(data)), ]
} else {
  stop("A tabela não contém linhas suficientes para remover as primeiras 6 e as últimas 2 linhas.")
}

# Remover a última coluna da tabela
data <- data[, -ncol(data)]

# Adicionar o cabeçalho especificado
colnames(data) <- c("Ano", "Data", "Produto", "Cliente", "Canal Distribuição", "Integralidade", "Apolices Emitidas", "Apolices Anuladas", "Apolices Vigentes")


#Algumas alterações
# 1º - Alterar as últimas 3 colunas para formato numérico com separador decimal como vírgula
data[, c("Apolices Emitidas", "Apolices Anuladas", "Apolices Vigentes")] <- lapply(data[, c("Apolices Emitidas", "Apolices Anuladas", "Apolices Vigentes")], function(col) {
  as.numeric(as.character(col))
})

# 2º - Alterar os valores da coluna "Produto" para remover "1-" e manter apenas o nome
data$Produto <- gsub("^\\d+-", "", data$Produto)
data$`Canal Distribuição` <- gsub("^\\d+-", "", data$`Canal Distribuição`)

# 3º - Adicionar uma nova coluna "num_mes" depois da coluna "Data"
data$Mes <- gsub("^\\d{4}/Jan$", "Janeiro",
                gsub("^\\d{4}/Feb$", "Fevereiro",
                gsub("^\\d{4}/Mar$", "Março",
                gsub("^\\d{4}/Apr$", "Abril",
                gsub("^\\d{4}/May$", "Maio",
                gsub("^\\d{4}/Jun$", "Junho",
                gsub("^\\d{4}/Jul$", "Julho",
                gsub("^\\d{4}/Aug$", "Agosto",
                gsub("^\\d{4}/Sep$", "Setembro",
                gsub("^\\d{4}/Oct$", "Outubro",
                gsub("^\\d{4}/Nov$", "Novembro",
                gsub("^\\d{4}/Dec$", "Dezembro", data$Data))))))))))))

# Reordenar as colunas 
data <- data[, c("Ano", "Data", "Produto", "Cliente", "Canal Distribuição", "Integralidade", "Apolices Emitidas", "Apolices Anuladas", "Apolices Vigentes")]

# Salvar o arquivo modificado
write_xlsx(data, output_file)

cat("O ficheiro foi modificado e guardado com sucesso em:", output_file, "\n")
