library(tidyr)
library(dplyr)
library(stringr)
library(openxlsx)

# This function add color to the simple logs
make_colorful_log <- function(log_file, questions, output_file = "./output/Master_logs.xlsx") {
  
  log <- read.xlsx(log_file, sheet = "02_Logbook")
  
  # Expand multiple question names and assign group_id
  question_name_to_group <- questions %>%
    mutate(group_id = row_number()) %>%
    separate_rows(question.name, sep = ",") %>%
    mutate(question.name = str_trim(question.name))
  
  # Join group_id and create group keys
  log <- log %>%
    left_join(question_name_to_group, by = "question.name") %>%
    mutate(group_key = paste(uuid, group_id, sep = "_")) %>%
    group_by(group_key) %>%
    mutate(group_number = cur_group_id()) %>%
    ungroup() %>%
    arrange(group_number)
  
  # Load workbook and modify existing file
  wb_old <- loadWorkbook(log_file)
  wb <- copyWorkbook(wb_old)
  sheets <- names(wb)
  if ("02_Logbook" %in% sheets) removeWorksheet(wb, "02_Logbook")
  
  addWorksheet(wb, "02_Logbook")
  writeData(wb, "02_Logbook", log)
  setColWidths(wb, "02_Logbook", cols = 1:(ncol(log) - 3), width = "auto")
  worksheetOrder(wb) <- match(sheets, names(wb))
  freezePane(wb, "02_Logbook", firstRow = TRUE)
  
  # Assign colors
  set.seed(12356)
  unique_groups <- unique(log$group_number)
  group_colors <- setNames(
    grDevices::rgb(runif(length(unique_groups), 0.4, 1),
                   runif(length(unique_groups), 0.4, 1),
                   runif(length(unique_groups), 0.4, 1)),
    unique_groups
  )
  
  # Apply color styles by group
  for (group in unique_groups) {
    rows <- which(log$group_number == group) + 1
    style <- createStyle(fgFill = group_colors[as.character(group)])
    addStyle(wb, "02_Logbook", style, rows = rows, cols = 1:(ncol(log) - 3), gridExpand = TRUE)
  }
  
  deleteData(wb, "02_Logbook", cols = (ncol(log) - 2):ncol(log), rows = 1:(nrow(log) + 1), gridExpand = TRUE)
  
  # Save changes
  saveWorkbook(wb, output_file, overwrite = TRUE)
}

