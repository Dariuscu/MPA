# Author: Dario Rodriguez
# Date started: 12/02/2015
# Date finalised: xxxx
# Scope: call some stored procedures in SQL Server that will produce different tables,
## then import these tables into R and group by Country and Region. Export the results to
## an Excel spreadhseets with 7 final tables
# --------------------------------------------------------------------------------------------------

# Clear any previous data
rm(list=ls())


# Time the script
startTime <- Sys.time()


# Load the library to work with SQL Server
library(RODBC)
library(data.table)
library(xlsx)


# Stop importing as factors
options(stringsAsFactors = FALSE)


# Set the working directory to the folder where I have the script
setwd("Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Scripts/MPA_Live_spreadsheets")


# Connect to SQL Server
myStringConn <- "Driver=SQL Server;Server=SQL-SPATIAL;Database=UKProtectedAreas;Trusted_Connection=True;"
conn <- odbcDriverConnect(myStringConn)


# Run stored procedures
sqlQuery(conn, "Offhore_SACs_SPAs_OffAndInshore")  #It will produce 9 tables


# Import the tables created through the stored procedure into a data frame for each
offshore_Sites <- sqlQuery(conn, "select * from Offshore_Sites")
offshore_Protected_Features <- sqlQuery(conn, "select * from Offshore_Protected_Features")
sacs_Sites <- sqlQuery(conn, "select * from SACs_Sites")
sacs_Marine_Interest_Features <- sqlQuery(conn, "select * from SACs_Marine_Interest_Features")
spas_Sites <- sqlQuery(conn, "select * from SPAs_Sites")
spas_Marine_Interest_Features <- sqlQuery(conn, "select * from SPAs_Marine_Interest_Features")
spas_Marine_Habitats <- sqlQuery(conn, "select * from SPAs_Marine_Habitats")
offshore_inshore_sites <- sqlQuery(conn, "select * from Offshore_and_Inshore_Sites")
offshore_inshore_protected_feat <- sqlQuery(conn, "select * from Offshore_and_Inshore_Protected_Features")


# Delete the tables since they have now been saved in data frames
sqlQuery(conn, "Delete_Offshore_SACs_SPAs_OffAndInshore")


# Close connection
odbcClose(conn)


# Group by Country and Region, concatenate and remove duplicates
# 1) Offshore sites
offshore_sites <- as.data.table(offshore_Sites)
offshore_sites <- setnames(offshore_sites, 1:10, c("code","name","status","country","region","area","long","lat","est_date","agency"))
#offshore_sites <- setkey(offshore_sites, code)
offshore_sites <- unique(offshore_sites[, list(code,
                                      name,
                                      status,
                                      country = paste(sort(unique(country)), collapse = ' & '),
                                      region = paste(sort(unique(region)), collapse = ' & '),
                                      area,
                                      long,
                                      lat,
                                      est_date,
                                      agency
                                      ), by = code
                               ])
offshore_sites <- offshore_sites[, 2:11, with = FALSE]

# 2) Offshore protected features
offshore_protected_features <- as.data.table(offshore_Protected_Features)
offshore_protected_features <- setnames(offshore_protected_features, 1:10, c("code","name","status","country","region","feat_broad","feat_spec","mar_feat","common_name","pop_type"))
offshore_protected_features <- unique(offshore_protected_features[, list(code,
                                                                  name,
                                                                  status,
                                                                  country = paste(sort(unique(country)), collapse = ' & '),
                                                                  region = paste(sort(unique(region)), collapse = ' & '),
                                                                  feat_broad,
                                                                  feat_spec,
                                                                  mar_feat,
                                                                  common_name,
                                                                  pop_type), by = code
                                                           ])                                
offshore_protected_features <- offshore_protected_features[, 2:11, with = FALSE]

# 3) SACs Sites
sacs_sites <- as.data.table(sacs_Sites)
sacs_sites <- setnames(sacs_sites, 1:13, c("code","name","status","country","region","area","long","lat","est_date","conf_date", "des_date", "agency", "non_marine"))
sacs_sites <- unique(sacs_sites[, list(code,
                                name,
                                status,
                                country = paste(sort(unique(country)), collapse = ' & '),
                                region = paste(sort(unique(region)), collapse = ' & '),
                                area,
                                long,
                                lat,
                                est_date,
                                conf_date,
                                des_date,
                                agency,
                                non_marine), by = code
                         ])
sacs_sites <- sacs_sites[, 2:14, with = FALSE]

# 4) SACs marine interest features
sacs_mar_int_fea <- as.data.table(sacs_Marine_Interest_Features)
sacs_mar_int_fea <- setnames(sacs_mar_int_fea, 1:8, c("code","name","status","country","region", "type","qualifying","common_name"))
sacs_mar_int_fea <- unique(sacs_mar_int_fea[, list(code,                          
                                            name,
                                            status,
                                            country = paste(sort(unique(country)), collapse = ' & '),
                                            region = paste(sort(unique(region)), collapse = ' & '),
                                            type,
                                            qualifying,
                                            common_name), by = code
                                      ])
sacs_mar_int_fea <- sacs_mar_int_fea[, 2:9, with = FALSE]

# 5) SPAs sites
spas_sites <- as.data.table(spas_Sites)
spas_sites <- setnames(spas_sites, 1:11, c("code","name","status","country","region","area","long","lat","first_date", "agency", "non_marine"))
spas_sites <- unique(spas_sites[, list(code,                          
                                name,
                                status,
                                country = paste(sort(unique(country)), collapse = ' & '),
                                region = paste(sort(unique(region)), collapse = ' & '),
                                area,
                                long,
                                lat,
                                first_date,
                                agency,
                                non_marine), by = code
                        ])
spas_sites <- spas_sites[, 2:12, with = FALSE]

# 6) SPAs marine interest features
spas_mar_int_fea <- as.data.table(spas_Marine_Interest_Features)
spas_mar_int_fea <- setnames(spas_mar_int_fea, 1:10, c("code","name","status","country","region","type","features","common_name","pop_type", "annex"))
spas_mar_int_fea <- unique(spas_mar_int_fea[, list(code,                          
                                            name,
                                            status,
                                            country = paste(sort(unique(country)), collapse = ' & '),
                                            region = paste(sort(unique(region)), collapse = ' & '),
                                            type,
                                            features,
                                            common_name,
                                            pop_type,
                                            annex),  by = code
                                     ])
spas_mar_int_fea <- spas_mar_int_fea[, 2:11, with = FALSE]

# 7) SPAs marine habitats
spas_mar_hab <- as.data.table(spas_Marine_Habitats)
spas_mar_hab <- setnames(spas_mar_hab, 1:6, c("code","name","status","country","region","habitats"))
spas_mar_hab <- unique(spas_mar_hab [, list(code,                          
                                            name,
                                            status,
                                            country = paste(sort(unique(country)), collapse = ' & '),
                                            region = paste(sort(unique(region)), collapse = ' & '),
                                            habitats),  by = code
                                     ])
spas_mar_hab  <- spas_mar_hab[, 2:7, with = FALSE]

# 8) Offshore and inshore sites
offin_sites <- as.data.table(offshore_inshore_sites)
offin_sites <- setnames(offin_sites, 1:13, c("code", "name", "status", "country", "region", "area", "long", "lat", "est_date", "agency", "precode", "wdpa", "nonmarine"))
offin_sites <- unique(offin_sites [, list(code,
                                          name,
                                          status,
                                          country = paste(sort(unique(country)), collapse = ' & '),
                                          region = paste(sort(unique(region)), collapse = ' & '),
                                          area,
                                          long,
                                          lat,
                                          est_date,
                                          agency,
                                          precode,
                                          wdpa,
                                          nonmarine), by = code
                                   ])
offin_sites <- offin_sites[, 2:14, with = FALSE]

# 9) Offshore and inshore protected features
offin_prot_fea <- as.data.table(offshore_inshore_protected_feat)
offin_prot_fea <- setnames(offin_prot_fea, 1:11, c("code", "name", "status", "country", "region", "feat_broad", "feat_spec", "feat_code", "mar_feat", "common_name","pop_type"))
offin_prot_fea <- unique(offin_prot_fea [, list(code,
                                          name,
                                          status,
                                          country = paste(sort(unique(country)), collapse = ' & '),
                                          region = paste(sort(unique(region)), collapse = ' & '),
                                          feat_broad,
                                          feat_spec,
                                          feat_code,
                                          mar_feat,
                                          common_name,
                                          pop_type), by = code
                                   ])
offin_prot_fea <- offin_prot_fea[, 2:12, with = FALSE]



# Formate the date fields
offshore_sites$est_date <- format(as.Date(offshore_sites$est_date, '%Y-%m-%d'), '%d-%m-%Y')
#sapply(offshore_sites, class)
sacs_sites$est_date <- format(as.Date(sacs_sites$est_date, '%Y-%m-%d'), '%d-%m-%Y')
sacs_sites$conf_date <- format(as.Date(sacs_sites$conf_date, '%Y-%m-%d'), '%d-%m-%Y')
sacs_sites$des_date <- format(as.Date(sacs_sites$des_date, '%Y-%m-%d'), '%d-%m-%Y')
spas_sites$first_date <- format(as.Date(spas_sites$first_date, '%Y-%m-%d'), '%d-%m-%Y')
offin_sites$est_date <- format(as.Date(offin_sites$est_date, '%Y-%m-%d'), '%d-%m-%Y')


# # Remove NAs from data tables
# offshore_sites[is.na(offshore_sites)] <- ""


# Put back the names of the fields
offshore_sites <- setnames(offshore_sites, 1:10, c("Site code","Site name","Site status","Country","CP2 Region","Area (ha)","Longitude","Latitude","Established date","Agency"))
offshore_protected_features <- setnames(offshore_protected_features, 1:10, c("Site code","Site name","Site status","Country","CP2 Region","Feature type (broad)","Feature type (specific)","Marine features protected","Common name","Population type"))
sacs_sites <- setnames(sacs_sites, 1:13, c("Site code","Site name","Site status","Country","CP2 Region","Area (ha)","Longitude","Latitude","cSAC Established Date","SCI Confirmation Date", "SAC Designation Date", "Agency", "Non-marine interest features"))
sacs_mar_int_fea <- setnames(sacs_mar_int_fea, 1:8, c("Site code","Site name","Site status","Country","CP2 Region", "Feature type","Qualifying marine interest features","Lay title or common name"))
spas_sites <- setnames(spas_sites, 1:11, c("Site code","Site name","Site status","Country","CP2 Region","Area (ha)","Longitude","Latitude","First Classification Date", "Agency", "Non-marine interest features"))
spas_mar_int_fea <- setnames(spas_mar_int_fea, 1:10, c("Site code","Site name","Site status","Country","CP2 Region","Feature type","Marine interest features","Common name","Population typee", "Annex I/ROMS"))
spas_mar_hab <- setnames(spas_mar_hab, 1:6, c("Site code","Site name","Site status","Country","CP2 Region","Habitats hosted"))
offin_sites <- setnames(offin_sites, 1:13, c("Site code", "Site name", "Site status", "Country", "CP2 Region", "Area (ha)", "Longitude", "Latitude", "Established date", "Agency", "Pre-Designation Code", "WDPA code", "Non-marine interest features"))
offin_prot_fea <- setnames(offin_prot_fea, 1:11, c("Site code", "Site name", "Site status", "Country", "CP2 Region", "Feature type (broad)", "Feature type (specific)", "Feature code", "Marine features protected", "Common name", "Population type"))


# Write all the data tables in an only excel spreadsheet
#write.table(sacs_sites, "C:/Users/Dario Rodriguez/Desktop/sites.csv", sep=",", row.names=FALSE, col.names=TRUE)
outwb <- createWorkbook()
saveWorkbook(outwb, "Offshore_SACs_SPAs_Offin.xlsx")
file <- "Offshore_SACs_SPAs_Offin.xlsx"

write.xlsx(offshore_sites, file, sheetName="Offshore Sites", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(offshore_protected_features, file, sheetName="Offshore Protected Features", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(sacs_sites, file, sheetName="SACs Sites", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(sacs_mar_int_fea, file, sheetName = "SACs Marine Interested Features", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(spas_sites, file, sheetName = "SPAs Sites", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(spas_mar_int_fea, file, sheetName="SPAs Marine Interested Features", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(spas_mar_hab, file, sheetName="SPAs Habitats", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(offin_sites, file, sheetName="OffAndIn Sites", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)
write.xlsx(offin_prot_fea, file, sheetName="OffAndIn Protected Features", col.names = TRUE, row.names = FALSE, append = TRUE, showNA = FALSE)


# Calculate the time the script took to run
round(Sys.time() - startTime, 1)