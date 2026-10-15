## Format O3 DV spreadsheets based on template file
o3.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("O3",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.1979 <- get.naa.info(par=44201,psid=9)
  naa.1997 <- get.naa.info(par=44201,psid=10)
  naa.2008 <- get.naa.info(par=44201,psid=11)
  naa.2015 <- get.naa.info(par=44201,psid=23)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=44201,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude",
    "naa_name_2015","naa_name_2008","naa_name_1997","naa_name_1979")]
  base_info$latitude <- as.numeric(base_info$latitude)
  base_info$longitude <- as.numeric(base_info$longitude)
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
  table0 <- data.frame(parameter="44201",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem=mco$frm_fem,non_ref=mco$non_ref,combo_site,combo_date,agency,
    collection_frequency=cfr,season_begin_date=sbd,season_end_date=sed,monitor_types=types,
    monitor_networks=nets,measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=12,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[12]]$sheet_data$rows))
  removeRowHeights(wb,sheet=12,rows=clear.rows)
  deleteData(wb,sheet=12,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=12,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=12,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))

  ## Pull 8-hour Ozone DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_OZONE_DVS
    WHERE design_value IS NOT NULL
      AND dv_year >= ",years[3],"
      AND dv_year <= ",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '44201'
      AND standard_id IN (10,11,23)
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year",sep=""),dbname="aqsprod")
  write.csv(subset(t,standard_id == 23),file=paste("DVs2xlsx/",year,"/O3dvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  v <- subset(t,standard_id == 10)
  t1997 <- data.frame(site=paste(v$state_code,v$county_code,v$site_number,sep=""),
    dv_year=as.integer(v$dv_year),dv=round(as.numeric(v$design_value),3),
    valid=v$dv_validity_indicator,pct_3yr=as.integer(v$dv_3_yr_percent_complete),
    pct_yr1=as.integer(v$year_2_percent_complete),pct_yr2=as.integer(v$year_1_percent_complete),
    pct_yr3=as.integer(v$dv_year_percent_complete),max4_yr1=round(as.numeric(v$year_2_4th_max),3),
    max4_yr2=round(as.numeric(v$year_1_4th_max),3),max4_yr3=round(as.numeric(v$dv_year_4th_max),3))
  sites <- table0[!duplicated(table0$site),c("site","epa_region","state_name","county_name",
    "cbsa_name","csa_name","naa_name_1997","site_name","address","latitude","longitude")]
  dv.1997 <- merge(sites,t1997,by="site")
  v <- subset(t,standard_id == 11)
  t2008 <- data.frame(site=paste(v$state_code,v$county_code,v$site_number,sep=""),
    dv_year=as.integer(v$dv_year),dv=round(as.numeric(v$design_value),3),
    valid=v$dv_validity_indicator,pct_3yr=as.integer(v$dv_3_yr_percent_complete),
    pct_yr1=as.integer(v$year_2_percent_complete),pct_yr2=as.integer(v$year_1_percent_complete),
    pct_yr3=as.integer(v$dv_year_percent_complete),max4_yr1=round(as.numeric(v$year_2_4th_max),3),
    max4_yr2=round(as.numeric(v$year_1_4th_max),3),max4_yr3=round(as.numeric(v$dv_year_4th_max),3))
  sites <- table0[!duplicated(table0$site),c("site","epa_region","state_name","county_name",
    "cbsa_name","csa_name","naa_name_2008","site_name","address","latitude","longitude")]
  dv.2008 <- merge(sites,t2008,by="site")
  v <- subset(t,standard_id == 23)
  t2015 <- data.frame(site=paste(v$state_code,v$county_code,v$site_number,sep=""),
    dv_year=as.integer(v$dv_year),dv=round(as.numeric(v$design_value),3),
    valid=v$dv_validity_indicator,pct_3yr=as.integer(v$dv_3_yr_percent_complete),
    pct_yr1=as.integer(v$year_2_percent_complete),pct_yr2=as.integer(v$year_1_percent_complete),
    pct_yr3=as.integer(v$dv_year_percent_complete),max4_yr1=round(as.numeric(v$year_2_4th_max),3),
    max4_yr2=round(as.numeric(v$year_1_4th_max),3),max4_yr3=round(as.numeric(v$dv_year_4th_max),3))
  sites <- table0[!duplicated(table0$site),c("site","epa_region","state_name","county_name",
    "cbsa_name","csa_name","naa_name_2015","site_name","address","latitude","longitude")]
  dv.2015 <- merge(sites,t2015,by="site")

  ## Calculate 1-hour ozone DVs from the raw data (most recent 3-year period only)
  source("R/get_rawdata.r")
  source(paste("OzoneDV/DV",year-2,"_",year,"/code/dvs_step1.r",sep=""))
  for (y in (year-2):year) { get.raw.data(par=44201,year=y) }
  design.values(years=(year-2):year,method="H",level=124,nonreg=FALSE,ee.type="C")

  ## Table 1a: 2015 8-hour NAAQS nonattainment area status
  temp <- subset(dv.2015,naa_name_2015 != " " & dv_year == year & valid == "Y")
  vals <- tapply(temp$dv,list(temp$naa_name_2015),max.na)
  vals <- merge(naa.2015,data.frame(naa_name=names(vals),dv=vals),by="naa_name",all=TRUE)
  table1a <- data.frame(naa.2015[,c("naa_name","epa_regions","status","classification")],dv=vals$dv,
    meets_naaqs=NA,naa.2015[,c("cdd_date","redesignation_date")],row.names=c(1:nrow(naa.2015)))
  table1a$meets_naaqs <- sapply(table1a$dv,function(x) ifelse(is.na(x),"Incomplete",
    ifelse(x > 0.070,"No","Yes")))
  writeData(wb,sheet=1,x=table1a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 1b: 2008 8-hour NAAQS nonattainment area status
  temp <- subset(dv.2008,naa_name_2008 != " " & dv_year == year & valid == "Y")
  vals <- tapply(temp$dv,list(temp$naa_name_2008),max.na)
  vals <- merge(naa.2008,data.frame(naa_name=names(vals),dv=vals),by="naa_name",all=TRUE)
  table1b <- data.frame(naa.2008[,c("naa_name","epa_regions","status","classification")],dv=vals$dv,
    meets_naaqs=NA,naa.2008[,c("cdd_date","redesignation_date")],row.names=c(1:nrow(naa.2008)))
  table1b$meets_naaqs <- sapply(table1b$dv,function(x) ifelse(is.na(x),"Incomplete",
    ifelse(x > 0.075,"No","Yes")))
  writeData(wb,sheet=2,x=table1b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  
  ## Table 1c: 1997 8-hour NAAQS nonattainment area status
  temp <- subset(dv.1997,naa_name_1997 != " " & dv_year == year & valid == "Y")
  vals <- tapply(temp$dv,list(temp$naa_name),max.na)
  vals <- merge(naa.1997,data.frame(naa_name=names(vals),dv=vals),by="naa_name",all=TRUE)
  table1c <- data.frame(naa.1997[,c("naa_name","epa_regions","status","classification")],dv=vals$dv,
    meets_naaqs=NA,naa.1997[,c("cdd_date","redesignation_date")],row.names=c(1:nrow(naa.1997)))
  table1c$meets_naaqs <- sapply(table1c$dv,function(x) ifelse(is.na(x),"Incomplete",
    ifelse(x > 0.084,"No","Yes")))
  writeData(wb,sheet=3,x=table1c,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 1d: 1979 1-hour NAAQS nonattainment area status
  load(paste("OzoneDV/DV",year-2,"_",year,"/",toupper(format(Sys.Date(),"%d%b%y")),
    "/dv_",year-2,"_",year,"_HRC124.Rdata",sep=""))
  temp <- subset(out,naa_name_1979 != " ",c("naa_name_1979","epa_region",paste("dv",year-2,year,sep="."),
    paste("ee",year-2,year,sep=".")))
  regions <- unlist(tapply(temp$epa_region,list(temp$naa_name_1979),
    function(x) paste(unique(x[order(x)]),collapse=",")))
  vals <- tapply(temp[,paste("dv",year-2,year,sep=".")],list(temp$naa_name_1979),max.na)
  exc <- tapply(temp[,paste("ee",year-2,year,sep=".")],list(temp$naa_name_1979),max.na)
  meets.naaqs <- sapply(exc,function(x) ifelse(is.na(x),"Incomplete",ifelse(x <= 1,"Yes","No")))
  table1d <- data.frame(naa_name=names(regions),region=regions,dv=vals/1000,
    ee=exc,meets_naaqs=meets.naaqs,row.names=c(1:length(regions)))
  writeData(wb,sheet=4,x=table1d,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  
  ## Table 2: Violating sites outside 2015 8-hour NAAQS nonattainment areas
  table2 <- subset(dv.2015,dv_year == year & naa_name_2015 == " " & dv > 0.070,c("state_name",
    "county_name","epa_region","site","dv","cbsa_name"))
  writeData(wb,sheet=5,x=table2,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 3a: Trends in 2015 8-hour NAAQS nonattainment areas
  temp <- subset(dv.2015,naa_name_2015 != " " & valid == "Y",c("naa_name_2015","dv_year","dv"))
  table3a <- naa.2015[,c("naa_name","epa_regions")]
  for (y in years[3:length(years)]) {
    table3a[,paste("dv",y-2,y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_2015","dv"))
    for (i in 1:nrow(naa.2015)) {
      v <- subset(t,naa_name_2015 == naa.2015$naa_name[i])
      if (nrow(v) == 0) { next }
      table3a[i,paste("dv",y-2,y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=6,x=table3a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=6,years); set.footnote.dates(wb,sheet=6);
  
  ## Table 3b: Trends in 2008 8-hour NAAQS nonattainment areas
  temp <- subset(dv.2008,naa_name_2008 != " " & valid == "Y",c("naa_name_2008","dv_year","dv"))
  table3b <- naa.2008[,c("naa_name","epa_regions")]
  for (y in years[3:length(years)]) {
    table3b[,paste("dv",y-2,y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_2008","dv"))
    for (i in 1:nrow(naa.2008)) {
      v <- subset(t,naa_name_2008 == naa.2008$naa_name[i])
      if (nrow(v) == 0) { next }
      table3b[i,paste("dv",y-2,y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=7,x=table3b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=7,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=7,years); set.footnote.dates(wb,sheet=7);
  
  ## Table 3c: Trends in 1997 8-hour NAAQS nonattainment areas
  temp <- subset(dv.1997,naa_name_1997 != " " & valid == "Y",c("naa_name_1997","dv_year","dv"))
  table3c <- naa.1997[,c("naa_name","epa_regions")]
  for (y in years[3:length(years)]) {
    table3c[,paste("dv",y-2,y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_1997","dv"))
    for (i in 1:nrow(naa.1997)) {
      v <- subset(t,naa_name_1997 == naa.1997$naa_name[i])
      if (nrow(v) == 0) { next }
      table3c[i,paste("dv",y-2,y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=8,x=table3c,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=8,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=8,years); set.footnote.dates(wb,sheet=8);
  
  ## Table 4: County-level design values for the 2015 8-hour NAAQS
  t <- subset(dv.2015,dv_year == year & valid == "Y",c("site",
    "state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4 <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$site,1,2),county_fips=substr(t$site,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv,cbsa_name=t$cbsa_name)
  table4 <- table4[order(table4$state_name,table4$county_name),]
  writeData(wb,sheet=9,x=table4,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=9,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=9,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=9,years); set.footnote.dates(wb,sheet=9);
  remove.extra.rows(wb,sheet=9,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5: Site-level design values for the 2015 8-hour NAAQS
  t <- subset(dv.2015,dv_year == year & !is.na(pct_3yr))
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv,t$valid)
  table5 <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2015","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv","pct_3yr",
    paste("pct_yr",1:3,sep=""),paste("max4_yr",1:3,sep=""))]
  writeData(wb,sheet=10,x=table5,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=10,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=10,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=10,years); set.footnote.dates(wb,sheet=10);
  remove.extra.rows(wb,sheet=10,row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 6: Trends in site-level design values for the 2015 8-hour NAAQS
  t <- dcast(subset(dv.2015,valid == "Y"),site ~ dv_year,value.var="dv")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dv.2015[,1:11],!duplicated(site)),t,by="site")
  table6 <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2015","epa_region",
    "site","site_name","address","latitude","longitude",paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=11,x=table6,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=11,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=11,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=11,years); set.footnote.dates(wb,sheet=11);
  remove.extra.rows(wb,sheet=11,row.hts=c(15,31,15,47,15,79,15))
  
  ## Write DV tables to Excel file
  file.xlsx <- paste("O3_DesignValues",year-2,year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}

