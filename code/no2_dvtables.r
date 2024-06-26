## Format NO2 DV spreadsheets based on template file
no2.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
  
  ## Load template DV spreadsheet file
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("NO2",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42602,psid=8)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42602,yr1=years[1],yr2=year,all=TRUE)
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
  table0 <- data.frame(parameter="42602",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,collection_frequency=cfr,
    season_begin_date=sbd,season_end_date=sed,monitor_types=types,monitor_networks=nets,
    measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=11,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[11]]$sheet_data$rows))
  removeRowHeights(wb,sheet=11,rows=clear.rows)
  deleteData(wb,sheet=11,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=11,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=11,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))
  
  ## Pull annual and 1-hour NO2 DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_NO2_ANNUAL_DVS
    WHERE design_value IS NOT NULL
      AND dv_year >= ",years[3],"
      AND dv_year <= ",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '42602'
      AND pollutant_standard_id = 8
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year",sep=""))
  write.csv(t,file=paste("DVs2xlsx/",year,"/NO2dvann",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.ann <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv_ann=round(as.numeric(t$design_value)),
    pct_ann=as.integer(t$observation_percent),valid_ann=t$dv_validity_indicator)
  dvs.ann <- dvs.ann[order(dvs.ann$site,dvs.ann$dv_year,dvs.ann$pct_ann,decreasing=TRUE),]
  dvs.ann <- dvs.ann[which(!duplicated(dvs.ann[,c("site","dv_year")])),]
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_NO2_1HOUR_DVS
    WHERE design_value > 0
      AND dv_year >= ",years[3],"
      AND dv_year <= ",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '42602'
      AND pollutant_standard_id = 20
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year",sep=""))
  write.csv(t,file=paste("DVs2xlsx/",year,"/NO2dv1hr",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.1hr <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv_1hr=as.numeric(t$design_value),valid_1hr=t$dv_validity_indicator,
    qtrs_yr1=as.integer(t$year_2_complete_quarters),qtrs_yr2=as.integer(t$year_1_complete_quarters),
    qtrs_yr3=as.integer(t$year_0_complete_quarters),p98_yr1=as.numeric(t$year_2_98th_percentile),
    p98_yr2=as.numeric(t$year_1_98th_percentile),p98_yr3=as.numeric(t$year_0_98th_percentile))
  sites <- table0[!duplicated(table0$site),c("site","epa_region","state_name","county_name",
    "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude")]
  dvs <- merge(sites,merge(dvs.ann,dvs.1hr,by=c("site","dv_year"),all=TRUE),by="site")
  dvs <- dvs[order(dvs$site,dvs$dv_year),]
  
  ## Table 1a: Nonattainment area status for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,naa_name != " " & dv_year == year & valid_ann == "Y",c("naa_name","dv_ann"))
  t <- subset(t[order(t$naa_name,t$dv_ann,decreasing=TRUE),],!duplicated(naa_name))
  table1a <- merge(naa.info[,c("naa_name","epa_regions","status")],t,by="naa_name",all=TRUE)
  table1a$meets_naaqs <- sapply(table1a$dv,function(x) ifelse(x > 53,"No","Yes"))
  writeData(wb,sheet=1,x=table1a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 2a: Additional monitors violating the 1971 Annual NO2 NAAQS
  table2a <- subset(dvs,naa_name != " " & dv_ann > 53 & valid_ann == "Y",
    c("state_name","county_name","epa_region","site","dv_ann","cbsa_name"))
  if (nrow(table2a) == 0) {
    table2a <- data.frame(x=paste("There were no sites violating the annual NO2 NAAQS in ",
      year,".",sep=""))
  }
  writeData(wb,sheet=2,x=table2a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=2,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  
  ## Table 2b: Monitors violating the 2010 1-hour NO2 NAAQS
  table2b <- subset(dvs,dv_1hr > 100 & valid_1hr == "Y", 
    c("state_name","county_name","epa_region","site","dv_1hr","cbsa_name"))
  if (nrow(table2b) == 0) {
    table2b <- data.frame(x=paste("There were no sites violating the 1-hour NO2 NAAQS in ",
      year-2,"-",year,".",sep=""))
  }
  writeData(wb,sheet=3,x=table2b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=3,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 3a: Nonattainment area trends for the 1971 Annual NO2 NAAQS
  temp <- subset(dvs,naa_name != " " & valid_ann == "Y",c("naa_name","dv_year","dv_ann"))
  table3a <- naa.info[,c("naa_name","epa_regions")]
  for (y in years[3:length(years)]) {
    table3a[,paste("dv",y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv_ann"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3a[i,paste("dv",y,sep="_")] <- max(v$dv_ann)
    }
  }
  writeData(wb,sheet=4,x=table3a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  
  ## Table 4a: County-level design values for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,dv_year == year & valid_ann == "Y",c("site",
    "state_name","county_name","epa_region","dv_ann","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$site,1,2),county_fips=substr(t$site,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv_ann,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  writeData(wb,sheet=5,x=table4a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 4b: County-level design values for the 2010 1-hour NO2 NAAQS
  t <- subset(dvs,dv_year == year & valid_1hr == "Y",c("site",
    "state_name","county_name","epa_region","dv_1hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_1hr,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$site,1,2),county_fips=substr(t$site,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv_1hr,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  writeData(wb,sheet=6,x=table4b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=6,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=6,years,row=4); set.footnote.dates(wb,sheet=6);
  remove.extra.rows(wb,sheet=6,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5a: Site-level design values for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,dv_year == year & pct_ann > 0)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv_ann,t$valid_ann)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv_ann,t$valid_ann)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv","pct_ann")]
  writeData(wb,sheet=7,x=table5a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=7,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=7,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=7,years); set.footnote.dates(wb,sheet=7);
  remove.extra.rows(wb,sheet=7,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 5b: Site-level design values for the 2010 1-hour NO2 NAAQS
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv_1hr,t$valid_1hr)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv_1hr,t$valid_1hr)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    "qtrs_yr1","qtrs_yr2","qtrs_yr3","p98_yr1","p98_yr2","p98_yr3")]
  writeData(wb,sheet=8,x=table5b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=8,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=8,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=8,years); set.footnote.dates(wb,sheet=8);
  remove.extra.rows(wb,sheet=8,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6a: Trends in site-level design values for the 1971 Annual NO2 NAAQS
  t <- dcast(subset(dvs,valid_ann == "Y"),site ~ dv_year,value.var="dv_ann")
  colnames(t)[2:11] <- paste("dv",years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6a <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude",paste("dv",years[3:12],sep="_"))]
  writeData(wb,sheet=9,x=table6a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=9,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=9,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=9,years); set.footnote.dates(wb,sheet=9);
  remove.extra.rows(wb,sheet=9,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6b: Trends in site-level design values for the 2010 1-hour NO2 NAAQS
  t <- dcast(subset(dvs,valid_1hr == "Y"),site ~ dv_year,value.var="dv_1hr")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6b <- vals[,c("state_name","county_name","cbsa_name","csa_name","epa_region","site",
    "site_name","address","latitude","longitude",paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=10,x=table6b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=10,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=10,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=10,years); set.footnote.dates(wb,sheet=10);
  remove.extra.rows(wb,sheet=10,row.hts=c(15,31,15,47,15,79,15))
  
  ## Save DV tables in .Rdata format and write to Excel file
  file.xlsx <- paste("NO2_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}
