## Format CO DV spreadsheets based on template file
co.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
  type=ifelse(as.numeric(format(Sys.Date(),"%m")) < 5,"DRAFT","FINAL")) {
  
  ## Custom functions called within main function
  max.na <- function(x) ifelse(any(!is.na(x)),max(x,na.rm=TRUE),NA)
  met.naaqs <- function(x) {
    y <- data.frame(A=sum(x == "A"),I=sum(x == "I"),V=sum(x == "V"))
    z <- ifelse(y$V > 0,"No",ifelse(y$A == 0,"Incomplete","Yes"))
    return(z)
  }
  get.naa.states <- function(df) {
    naa.names <- df$naa_name
    states <- unlist(sapply(naa.names,function(x) strsplit(x,split=", ")[[1]][2]))
    states <- unlist(sapply(states,function(x) strsplit(x,split="-")))
    codes <- unique(states[order(states)])
    state.codes <- c(state.abb,"DC","PR")
    state.names <- c(state.name,"District of Columbia","Puerto Rico")
    out <- state.names[match(codes,state.codes)]
    return(out)
  }
  
  ## Set up working environment
  source("C:/Users/bwells01/Documents/R/get_monitors.r")
  source("C:/Users/bwells01/Documents/R/openxlsx_dvfuns.r")
  require(openxlsx); require(plyr); require(reshape2); 
  dv.begin.date <- as.Date(paste(year-1,"01-01",sep="-"))
  today <- format(Sys.Date(),"%x")
  years <- c((year-10):year)
  
  ## Load template DV spreadsheet file
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("CO",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42101,psid=4)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42101,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_1971")]
  base_info$latitude <- as.numeric(base_info$latitude)
  base_info$longitude <- as.numeric(base_info$longitude)
  colnames(base_info)[ncol(base_info)] <- "naa_name"
  mbd <- get.unique.dates(monitors$monitor_begin_date,monitors$id,first=TRUE)
  med <- get.unique.dates(monitors$monitor_end_date,monitors$id,first=FALSE)
  lsd <- get.unique.dates(monitors$last_sample_date,monitors$id,first=FALSE)
  pbd <- get.unique.dates(monitors$primary_begin_date,monitors$id,first=TRUE)
  ped <- get.unique.dates(monitors$primary_end_date,monitors$id,first=FALSE)
  sbd <- substr(get.unique.codes(monitors$season_begin_date,monitors$id,monitors$season_begin_date),6,10)
  sed <- substr(get.unique.codes(monitors$season_begin_date,monitors$id,monitors$season_end_date),6,10)
  nbd <- get.unique.dates(monitors$nonreg_begin_date,monitors$id,first=TRUE)
  ned <- get.unique.dates(monitors$nonreg_end_date,monitors$id,first=FALSE)
  nrc <- get.unique.codes(monitors$nonreg_begin_date,monitors$id,monitors$nonreg_concur)
  mtc <- get.unique.codes(monitors$method_begin_date,monitors$id,monitors$method_code)
  msc <- get.unique.codes(monitors$monitor_begin_date,monitors$id,monitors$measurement_scale)
  obj <- get.unique.codes(monitors$monitor_begin_date,monitors$id,monitors$monitor_objective)
  cfr <- get.unique.codes(monitors$monitor_begin_date,monitors$id,monitors$collection_frequency)
  frm_fem <- non_ref <- rep(" ",length(ids)); check.dates <- rep(FALSE,length(ids));
  for (i in 1:length(ids)) {
    codes <- subset(monitors,id == ids[i],
      c("method_code","frm_code","method_begin_date","method_end_date"))
    codes <- codes[which(!duplicated(codes)),]
    codes <- codes[order(codes$method_code),]
    if (sum(codes$frm_code != " ") > 0) { 
      frm_fem[i] <- paste(unique(codes$method_code[which(codes$frm_code != " ")]),collapse=",")
    }
    if (sum(codes$frm_code == " ") > 0) { 
      non_ref[i] <- paste(unique(codes$method_code[which(codes$frm_code == " ")]),collapse=",")
    }
    if (non_ref[i] != " ") {
      nr <- subset(codes,frm_code == " ")
      nr$method_end_date <- sapply(nr$method_end_date,function(x) ifelse(x == " ",Sys.Date(),as.Date(x)))
      if (any(nr$method_end_date > dv.begin.date)) { check.dates[i] <- TRUE }
    }
  }
  monitors$monitor_type <- gsub("NON-REGULATORY"," ",gsub("OTHER"," ",monitors$monitor_type))
  types <- gsub(" ,","",tapply(monitors$monitor_type,list(monitors$id),function(x) 
    paste(unique(x[order(x)]),collapse=",")))
  if (any(types == " ")) { types[which(types == " ")] <- "OTHER" }
  monitors$network <- gsub("UNOFFICIAL ","",gsub("PROPOSED ","",monitors$network))
  nets <- tapply(monitors$network,list(monitors$id),function(x)
    paste(unique(x[order(x)]),collapse=","))
  combos <- subset(monitors,combo_site != " ",c("id","combo_site","combo_date"))
  combos <- combos[which(!duplicated(combos)),]
  combo_site <- sapply(substr(ids,1,9),function(x) ifelse(x %in% substr(combos$id,1,9),
    combos$combo_site[match(x,substr(combos$id,1,9))],ifelse(x %in% combos$combo_site,
    substr(combos$id[match(x,combos$combo_site)],1,9)," ")))
  combo_date <- sapply(substr(ids,1,9),function(x) ifelse(x %in% substr(combos$id,1,9),
    combos$combo_date[match(x,substr(combos$id,1,9))],ifelse(x %in% combos$combo_site,
    combos$combo_date[match(x,combos$combo_site)]," ")))
  table0 <- data.frame(parameter="42101",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,collection_frequency=cfr,
    season_begin_date=sbd,season_end_date=sed,monitor_types=types,monitor_networks=nets,
    measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=13,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[13]]$sheet_data$rows))
  removeRowHeights(wb,sheet=13,rows=clear.rows)
  deleteData(wb,sheet=13,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=13,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=13,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))
  
  ## Pull CO DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_CO_DVS
    WHERE parameter_code = '42101'
      AND dv_year >=",years[2],
     "AND dv_year <=",year,
     "AND edt_id IN (0,5)
      AND si_id != 92355
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, poc, dv_year",sep=""))
  dvs.co <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    poc=t$poc,dv_year=as.integer(t$dv_year),
    dv_1hr=as.numeric(t$co_1hr_2nd_max_value),dt_1hr=as.character(t$co_1hr_2nd_max_date_time),
    dv_8hr=as.numeric(t$co_8hr_2nd_max_value),dt_8hr=as.character(t$co_8hr_2nd_max_date_time))   
  sites <- table0[,c("site","poc","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude")]
  dvs <- merge(sites,dvs.co,by=c("site","poc"))
  dvs$valid <- TRUE
  spm.check <- paste(table0$site,table0$poc,sep="")[which(grepl("SPM",table0$monitor_types) & 
    (as.Date(gsub(" ",Sys.Date(),table0$monitor_end_date)) - as.Date(table0$monitor_begin_date) <= 730 |
     as.Date(table0$last_sample_date) - as.Date(table0$monitor_begin_date) <= 730))]
  dvs$valid[which(paste(dvs$site,dvs$poc,sep="") %in% spm.check)] <- FALSE
  t$dv_validity_ind <- sapply(dvs$valid,function(x) ifelse(x,"Y","N"))
  write.csv(t,file=paste("DVs2xlsx/",year,"/COdvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  
  ## Table 1a. NAA Status 8hr
  t <- subset(dvs,naa_name != " " & dv_year == year & valid,c("naa_name","dv_8hr"))
  t <- subset(t[order(t$naa_name,t$dv_8hr,decreasing=TRUE),],!duplicated(naa_name))
  table1a <- merge(naa.info,t,by="naa_name",all=TRUE)[,c("naa_name","epa_regions","status","dv_8hr")]
  table1a$met_naaqs <- sapply(table1a$dv_8hr,function(x) ifelse(is.na(x),"No Data",
    ifelse(x > 9,"No","Yes")))
  writeData(wb,sheet=1,x=table1a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years,co=TRUE); set.footnote.dates(wb,sheet=1);
  
  ## Table 1b. NAA Status 1hr
  t <- subset(dvs,naa_name != " " & dv_year == year & valid,c("naa_name","dv_1hr"))
  t <- subset(t[order(t$naa_name,t$dv_1hr,decreasing=TRUE),],!duplicated(naa_name))
  table1b <- merge(naa.info,t,by="naa_name",all=TRUE)[,c("naa_name","epa_regions","status","dv_1hr")]
  table1b$met_naaqs <- sapply(table1b$dv_1hr,function(x) ifelse(is.na(x),"No Data",
    ifelse(x > 35,"No","Yes")))
  writeData(wb,sheet=2,x=table1b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=2,years,co=TRUE); set.footnote.dates(wb,sheet=2);
  
  ## Table 2a. Other Violators 8hr
  t <- subset(dvs,naa_name == " " & dv_year == year & dv_8hr > 9)
  if (nrow(t) == 0) {
    if (sum(table1a$met_naaqs == "No") == 0) {
      table2a <- data.frame(x=paste("There were no sites violating the 8-hour CO NAAQS in ",
        (year-1),"-",year,".",sep=""))
    }
    if (sum(table1a$met_naaqs == "No") > 0) {
      table2a <- data.frame(x=paste("There were no sites outside of areas previously designated",
       " nonattainment violating the 8-hour CO NAAQS in ",(year-1),"-",year,".",sep=""))
    } 
  }
  if (nrow(t) > 0) {
    table2a <- t[,c("state_name","county_name","epa_region","site","poc","dv_8hr","cbsa_name")]
  }
  writeData(wb,sheet=3,x=table2a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=3,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=3,years,co=TRUE); set.footnote.dates(wb,sheet=3);
  
  ## Table 2b. Other Violators 1hr
  t <- subset(dvs,naa_name == " " & dv_year == year & dv_1hr > 35)
  if (nrow(t) == 0) {
   if (sum(table1b$met_naaqs == "No") == 0) {
      table2b <- data.frame(x=paste("There were no sites violating the 1-hour CO NAAQS in ",
        (year-1),"-",year,".",sep=""))
    }
    if (sum(table1b$met_naaqs == "No") > 0) {
      table2b <- data.frame(x=paste("There were no sites outside of areas previously designated",
       " nonattainment violating the 1-hour CO NAAQS in ",(year-1),"-",year,".",sep=""))
    } 
  }
  if (nrow(t) > 0) {
    table2b <- t[,c("state_name","county_name","epa_region","site","poc","dv_1hr","cbsa_name")]
  }
  writeData(wb,sheet=4,x=table2b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=4,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=4,years,co=TRUE); set.footnote.dates(wb,sheet=4);
  
  ## Table 3a. NAA Trends 8hr
  temp <- subset(dvs,naa_name != " " & valid)
  table3a <- data.frame(naa_name=naa.info$naa_name,epa_regions=naa.info$epa_regions)
  for (y in years[2:length(years)]) {
    table3a[,paste("dv",(y-1),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv_8hr"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3a[i,paste("dv",(y-1),y,sep="_")] <- max(v$dv_8hr)
    }
  }
  writeData(wb,sheet=5,x=table3a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=5,years,co=TRUE); set.footnote.dates(wb,sheet=5);
  
  ## Table 3b. NAA Trends 1hr
  temp <- subset(dvs,naa_name != " " & valid)
  table3b <- data.frame(naa_name=naa.info$naa_name,epa_regions=naa.info$epa_regions)
  for (y in years[2:length(years)]) {
    table3b[,paste("dv",(y-1),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv_1hr"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3b[i,paste("dv",(y-1),y,sep="_")] <- max(v$dv_1hr)
    }
  }
  writeData(wb,sheet=6,x=table3b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=6,years,co=TRUE); set.footnote.dates(wb,sheet=6);
  
  ## Table 4a. County Status 8hr
  t <- subset(dvs,dv_year == year & valid,c("site","poc","state_name","county_name",
    "epa_region","dv_8hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_8hr,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,poc=t$poc,dv=t$dv_8hr,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  writeData(wb,sheet=7,x=table4a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=7,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=7,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=7,years,co=TRUE); set.footnote.dates(wb,sheet=7);
  remove.extra.rows(wb,sheet=7,row.hts=c(15,63,15,63,15,95,15))
  
  ## Table 4b. County Status 1hr
  t <- subset(dvs,dv_year == year & valid,c("site","poc","state_name","county_name",
    "epa_region","dv_1hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_1hr,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,poc=t$poc,dv=t$dv_1hr,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  writeData(wb,sheet=8,x=table4b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=8,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=8,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=8,years,co=TRUE); set.footnote.dates(wb,sheet=8);
  remove.extra.rows(wb,sheet=8,row.hts=c(15,63,15,63,15,95,15))
  
  ## Table 5a. Monitor Status 8hr
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid,dv,NA),t$dv_8hr,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid,NA,dv),t$dv_8hr,t$valid)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","poc","site_name","address","latitude","longitude","valid_dv","invalid_dv","dt_8hr")]
  writeData(wb,sheet=9,x=table5a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=9,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=9,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=9,years,co=TRUE); set.footnote.dates(wb,sheet=9);
  remove.extra.rows(wb,sheet=9,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 5b. Monitor Status 1hr
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid,dv,NA),t$dv_1hr,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid,NA,dv),t$dv_1hr,t$valid)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","poc","site_name","address","latitude","longitude","valid_dv","invalid_dv","dt_1hr")]
  writeData(wb,sheet=10,x=table5b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=10,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=10,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=10,years,co=TRUE); set.footnote.dates(wb,sheet=10);
  remove.extra.rows(wb,sheet=10,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6a. Monitor Trends 8hr
  t <- dcast(subset(dvs,valid),site + poc ~ dv_year,value.var="dv_8hr")
  colnames(t)[3:12] <- paste("dv",years[1:10],years[2:11],sep="_")
  table6a <- merge(sites,t,by=c("site","poc"))[,c("state_name","county_name","cbsa_name",
    "csa_name","naa_name","epa_region","site","poc","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[2:11],sep="_"))]
  writeData(wb,sheet=11,x=table6a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=11,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=11,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=11,years,co=TRUE); set.footnote.dates(wb,sheet=11);
  remove.extra.rows(wb,sheet=11,row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 6b. Monitor Trends 1hr
  t <- dcast(subset(dvs,valid),site + poc ~ dv_year,value.var="dv_1hr")
  colnames(t)[3:12] <- paste("dv",years[1:10],years[2:11],sep="_")
  table6b <- merge(sites,t,by=c("site","poc"))[,c("state_name","county_name","cbsa_name",
    "csa_name","naa_name","epa_region","site","poc","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[2:11],sep="_"))]
  writeData(wb,sheet=12,x=table6b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=12,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=12,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=12,years,co=TRUE); set.footnote.dates(wb,sheet=12);
  remove.extra.rows(wb,sheet=12,row.hts=c(15,47,15,47,15,79,15))
  
  ## Write DV tables to Excel File
  file.xlsx <- paste("CO_DesignValues",year-1,year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}