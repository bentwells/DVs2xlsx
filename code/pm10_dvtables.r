## Format PM10 DV spreadsheets based on template file
pm10.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
    states <- substr(naa.names,nchar(naa.names)-1,nchar(naa.names))
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
  dv.begin.date <- as.Date(paste(year-2,"01-01",sep="-"))
  today <- format(Sys.Date(),"%x")
  years <- c((year-11):year)
  
  ## Get DV input file and template file
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("PM10",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=81102,psid=12)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=81102,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_2006")]
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
  table0 <- data.frame(parameter="81102",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,collection_frequency=cfr,
    season_begin_date=sbd,season_end_date=sed,monitor_types=types,monitor_networks=nets,
    measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=7,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[7]]$sheet_data$rows))
  removeRowHeights(wb,sheet=7,rows=clear.rows)
  deleteData(wb,sheet=7,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=7,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=7,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))
  
  ## Pull PM10 DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_PM10_DVS
    WHERE  dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '81102'
      AND pollutant_standard_id = 12
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year"))
  write.csv(t,file=paste("DVs2xlsx/",year,"/PM10dvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.pm10 <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),poc=t$poc,
    dv_year=as.integer(t$dv_year),dv=as.numeric(t$dv_estimated_exceedances),valid=t$dv_validity_indicator,
    est.yr1=as.numeric(t$year_2_est_exceedances),est.yr2=as.numeric(t$year_1_est_exceedances),
    est.yr3=as.numeric(t$dv_year_est_exceedances),exc.yr1=as.numeric(t$year_2_exceedance_count),
    exc.yr2=as.numeric(t$year_1_exceedance_count),exc.yr3=as.numeric(t$dv_year_exceedance_count),
    qtrs.yr1=as.integer(t$year_2_complete_quarters),qtrs.yr2=as.integer(t$year_1_complete_quarters),
    qtrs.yr3=as.integer(t$dv_year_complete_quarters))
  sites <- subset(table0,!duplicated(site),c("site","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude"))
  dvs <- merge(sites,dvs.pm10)
  
  ## Table 1. NAA Status
  naa.test <- subset(dvs,naa_name != " " & dv_year == year)
  t <- subset(dvs,naa_name != " " & dv_year == year & valid == "Y",c("naa_name","dv"))
  t <- subset(t[order(t$naa_name,t$dv,decreasing=TRUE),],!duplicated(naa_name))
  table1 <- merge(naa.info[,c("naa_name","epa_regions","status")],t,by="naa_name",all=TRUE)
  table1$met_naaqs <- mapply(function(naa,dv) ifelse(naa %in% unique(naa.test$naa_name),
    ifelse(is.na(dv),"Incomplete",ifelse(dv > 1,"No","Yes")),"No Data"),table1$naa_name,table1$dv)
  writeData(wb,sheet=1,x=table1,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 2. Other Violators
  t <- subset(dvs,naa_name == " " & dv_year == year & dv > 1 & valid == "Y")
  table2 <- t[,c("state_name","county_name","epa_region","site","poc","dv","cbsa_name")]
  if (nrow(table2) == 0) {
    table2 <- data.frame(x=paste("There were no sites violating the PM10 NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  writeData(wb,sheet=2,x=table2,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=2,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  remove.extra.rows(wb,sheet=2,row.hts=c(15,33,15,47,15,63,15))
  
  ## Table 3. NAA Trends
  temp <- subset(dvs,naa_name != " " & valid == "Y")
  table3 <- data.frame(naa_name=naa.info$naa_name,epa_region=naa.info$epa_regions)
  for (y in years[3:length(years)]) {
    table3[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=3,x=table3,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 4. County Status
  t <- subset(dvs,dv_year == year & valid == "Y",
    c("site","poc","state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4 <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,poc=t$poc,dv=t$dv,cbsa_name=t$cbsa_name)
  table4 <- table4[order(table4$state_name,table4$county_name),]
  writeData(wb,sheet=4,x=table4,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=4,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  remove.extra.rows(wb,sheet=4,row.hts=c(15,49,15,63,15,95,15))
  
  ## Table 5. Monitor Status
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv,NA),t$dv,t$valid)
  table5 <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","poc","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    "est.yr1","est.yr2","est.yr3","exc.yr1","exc.yr2","exc.yr3","qtrs.yr1","qtrs.yr2","qtrs.yr3")]
  writeData(wb,sheet=5,x=table5,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,49,15,63,15,79,15))
  
  ## Table 6. Monitor Trends
  t <- dcast(subset(dvs,valid == "Y"),site + poc ~ dv_year,value.var="dv")
  colnames(t)[3:12] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(paste(site,poc))),t,by=c("site","poc"))
  table6 <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region","site","poc","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=6,x=table6,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=6,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=6,years); set.footnote.dates(wb,sheet=6);
  remove.extra.rows(wb,sheet=6,row.hts=c(15,33,15,63,15,63,15,79,15))
  
  ## Write DV tables to Excel File
  fix.scripts(wb)
  file.xlsx <- paste("PM10_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}