## Format SO2 DV spreadsheets based on template file
so2.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
  type=ifelse(as.numeric(format(Sys.Date(),"%m")) < 5,"DRAFT","FINAL")) {
  
  ## Set up working environment
  max.na <- function(x) ifelse(any(!is.na(x)),max(x,na.rm=TRUE),NA)
  source("R/get_monitors.r")
  source("R/openxlsx_dvfuns.r")
  require(openxlsx); require(plyr); require(reshape2);
  dv.begin.date <- as.Date(paste(year-2,"01-01",sep="-"))
  today <- format(Sys.Date(),"%x")
  years <- c((year-11):year)
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("SO2",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42401,psid=19)
  naa.info$type <- sapply(naa.info$designation_date,function(x) 
   ifelse(as.numeric(substr(gsub(" ",Sys.Date(),x),1,4)) %in% c(2016:2018,2025),"Modeled","Monitored"))
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42401,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_2010")]
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
  mco <- ddply(monitors,c("id"),summarize,frm_fem=ifelse(sum(frm_code != " ") == 0," ",
    paste(unique(method_code[which(frm_code != " ")]),collapse=",")),
    non_ref=ifelse(sum(frm_code == " ") == 0," ",
    paste(unique(method_code[which(frm_code == " ")]),collapse=",")))
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
  table0 <- data.frame(parameter="42401",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem=mco$frm_fem,non_ref=mco$non_ref,combo_site,combo_date,agency,
    collection_frequency=cfr,season_begin_date=sbd,season_end_date=sed,monitor_types=types,
    monitor_networks=nets,measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=10,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[10]]$sheet_data$rows))
  removeRowHeights(wb,sheet=10,rows=clear.rows)
  deleteData(wb,sheet=10,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=10,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=10,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))
  
  ## Pull 1-hour and annual DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_SO2_DVS
    WHERE design_value IS NOT NULL
      AND dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '42401'
      AND pollutant_standard_id = 19
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year"),dbname="aqsprod")
  write.csv(t,file=paste("DVs2xlsx/",year,"/SO2dvs1hr",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.1hr <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv.1hr=as.numeric(t$design_value),valid.1hr=t$dv_validity_indicator,
    ind.1hr.yr1=t$year_2_complete_ind,ind.1hr.yr2=t$year_1_complete_ind,ind.1hr.yr3=t$year_0_complete_ind,
    qtr.1hr.yr1=as.integer(t$year_2_complete_quarters),qtr.1hr.yr2=as.integer(t$year_1_complete_quarters),
    qtr.1hr.yr3=as.integer(t$year_0_complete_quarters),p99.yr1=as.numeric(t$year_2_99th_percentile),
    p99.yr2=as.numeric(t$year_1_99th_percentile),p99.yr3=as.numeric(t$year_0_99th_percentile))
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_SO2_DVS_SECONDARY
    WHERE design_value IS NOT NULL
      AND dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '42401'
      AND standard_id = 28
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_id, dv_year"),dbname="aqsprod")
  write.csv(t,file=paste("DVs2xlsx/",year,"/SO2dvsann",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.ann <- data.frame(site=paste(t$state_code,t$county_code,t$site_id,sep=""),
    dv_year=as.integer(t$dv_year),dv.ann=as.numeric(t$design_value),valid.ann=t$three_year_design_validity,
    ind.ann.yr1=t$year2_complete_ind,ind.ann.yr2=t$year1_complete_ind,ind.ann.yr3=t$year0_complete_ind,
    qtr.ann.yr1=as.integer(t$year2_complete_quarters),qtr.ann.yr2=as.integer(t$year1_complete_quarters),
    qtr.ann.yr3=as.integer(t$year0_complete_quarters),mean.yr1=as.numeric(t$year2_annual_mean),
    mean.yr2=as.numeric(t$year1_annual_mean),mean.yr3=as.numeric(t$year0_annual_mean))
  sites <- subset(table0,!duplicated(site),c("site","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude"))
  dvs <- merge(sites,merge(dvs.1hr,dvs.ann,by=c("site","dv_year"),all=TRUE),by="site")
  
  ## Table 1a. NAA Status 1-hour
  t <- subset(dvs,naa_name != " " & dv_year == year & valid.1hr == "Y",c("naa_name","dv.1hr"))
  t <- subset(t[order(t$naa_name,t$dv.1hr,decreasing=TRUE),],!duplicated(naa_name))
  table1a <- merge(naa.info,t,by="naa_name",all=TRUE)
  table1a$met_naaqs <- mapply(function(dv,type) ifelse(type == "Monitored",ifelse(is.na(dv),
    "Incomplete",ifelse(dv > 75,"No","Yes")),ifelse(is.na(dv),"[5]",ifelse(dv > 75,"No","[4]"))),
    dv=table1a$dv.1hr,type=table1a$type)
  monitored <- subset(table1a,type == "Monitored",c("naa_name","epa_regions","status",
    "dv.1hr","met_naaqs","cdd_date","redesignation_date"))
  modeled <- subset(table1a,type == "Modeled",c("naa_name","epa_regions","status",
    "dv.1hr","met_naaqs","cdd_date","redesignation_date"))
  writeData(wb,sheet=1,x=monitored,startCol=1,startRow=6,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=modeled,startCol=1,startRow=44,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 2a. Other Violators 1-hour
  t <- subset(dvs,naa_name == " " & dv_year == year & dv.1hr > 75 & valid.1hr == "Y")
  table2a <- t[,c("state_name","county_name","epa_region","site","dv.1hr","cbsa_name")]
  if (nrow(table2a) == 0) {
    table2a <- data.frame(x=paste("There were no sites violating the 1-hour SO2 NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  writeData(wb,sheet=2,x=table2a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=2,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  
  ## Table 3a. NAA Trends 1-hour
  temp <- subset(dvs,naa_name != " " & valid.1hr == "Y")
  table3a <- data.frame(naa_name=naa.info$naa_name,epa_regions=naa.info$epa_regions)
  for (y in years[3:length(years)]) {
    table3a[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv.1hr"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3a[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv.1hr)
    }
  }
  writeData(wb,sheet=3,x=table3a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 4a. County Status 1-hour
  t <- subset(dvs,dv_year == year & valid.1hr == "Y",
    c("site","state_name","county_name","epa_region","dv.1hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv.1hr,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv.1hr,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  writeData(wb,sheet=4,x=table4a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=4,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  remove.extra.rows(wb,sheet=4,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 4b. County Status Annual
  t <- subset(dvs,dv_year == year & valid.ann == "Y",
    c("site","state_name","county_name","epa_region","dv.ann","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv.ann,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv.ann,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  writeData(wb,sheet=5,x=table4b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5a. Site Status 1-hour
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv.1hr,t$valid.1hr)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv.1hr,t$valid.1hr)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region","site",
    "site_name","address","latitude","longitude","valid_dv","invalid_dv","ind.1hr.yr1","ind.1hr.yr2",
    "ind.1hr.yr3","qtr.1hr.yr1","qtr.1hr.yr2","qtr.1hr.yr3","p99.yr1","p99.yr2","p99.yr3")]
  writeData(wb,sheet=6,x=table5a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=6,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=6,years); set.footnote.dates(wb,sheet=6);
  remove.extra.rows(wb,sheet=6,row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 5b. Site Status Annual
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv.ann,t$valid.ann)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv.ann,t$valid.ann)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","epa_region","site","site_name",
    "address","latitude","longitude","valid_dv","invalid_dv","ind.ann.yr1","ind.ann.yr2",
    "ind.ann.yr3","qtr.ann.yr1","qtr.ann.yr2","qtr.ann.yr3","mean.yr1","mean.yr2","mean.yr3")]
  writeData(wb,sheet=7,x=table5b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=7,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=7,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=7,years); set.footnote.dates(wb,sheet=7);
  remove.extra.rows(wb,sheet=7,row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 6a. Site Trends 1-hour
  t <- dcast(subset(dvs,valid.1hr == "Y"),site ~ dv_year,value.var="dv.1hr")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6a <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=8,x=table6a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=8,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=8,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=8,years); set.footnote.dates(wb,sheet=8);
  remove.extra.rows(wb,sheet=8,row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6b. Site Trends Annual
  t <- dcast(subset(dvs,valid.ann == "Y"),site ~ dv_year,value.var="dv.ann")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6b <- vals[,c("state_name","county_name","cbsa_name","csa_name",
    "epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=9,x=table6b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=9,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=9,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=9,years); set.footnote.dates(wb,sheet=9);
  remove.extra.rows(wb,sheet=9,row.hts=c(15,31,15,47,15,79,15))
  
  ## Write DV tables to Excel File
  file.xlsx <- paste("SO2_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}