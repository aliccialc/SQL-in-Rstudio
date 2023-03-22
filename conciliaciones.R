#library(rJava)
#library(xlsxjars)
library(xlsx)
library(readxl)
library(sqldf)
library(RSQLite)

# table paths
# Windows
invoice= read_xlsx("Invoice.xlsx")
po=  read_xlsx("PO.xlsx")
report=  read_xlsx("Report.xlsx")

# Invoice

invoice_results = sqldf("SELECT i.`Código de cuenta`, r.`Rec. Reference`, i.`Resultado divisa`, r.Rights AS `Conciliacion ME`,
      CASE WHEN i.`Resultado divisa` < 0 THEN r.Rights + i.`Resultado divisa`
       ELSE i.`Resultado divisa` - r.Rights END AS diff
      FROM invoice AS i
      INNER JOIN report AS r
      ON r.`Rec. Reference` = i.Invoice
      WHERE i.`Código de cuenta` = 54201505")

#Purchase Orders

po_results = sqldf("SELECT p.`Código de cuenta`, p.PO, r.`Rec. Reference`, p.`Resultado divisa`,
      r.Rights AS `Conciliacion right`,
      
      CASE WHEN p.`Resultado divisa` < 0 THEN r.Rights + p.`Resultado divisa`
       ELSE p.`Resultado divisa` - r.Rights END AS `Difference right`,
       
       (r.Reserved + r.`Pending Fee`) AS `Conciliacion reverse`,
       
       CASE WHEN p.`Resultado divisa` < 0 THEN (r.Reserved + r.`Pending Fee`) + p.`Resultado divisa`
       ELSE p.`Resultado divisa` - (r.Reserved + r.`Pending Fee`) END AS `Difference reserved`
       
      FROM po AS p
      INNER JOIN report AS r
      WHERE p.`Código de cuenta` IN (54201005, 52101005, 54201905)
      AND r.Type = 'Purchase Order'
      AND r.Status = 'Advanced'")



# Path
my_path <- "C:\\Users\\" 


# Name of each sheet
data_names <- c("invoice_results", "po_results") 


#First Data Frame to a new excel file
write.xlsx(get(data_names[1]), paste0(my_path, "invoice_porders.xlsx"), row.names = FALSE, sheet = data_names[1])     

# for-loop to append each of the other data frames
for(i in 2:length(data_names)) {
  write.xlsx(get(data_names[i]), paste0(my_path, "invoice_porders.xlsx"), row.names = FALSE, sheetName = data_names[i], append = TRUE,
             ) 
}


