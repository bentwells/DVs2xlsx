## Format PM2.5 DV spreadsheets based on template file
pm25.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
  dv.begin.date <- as.Date(paste(year-2,"01-01",sep="-"))
  today <- format(Sys.Date(),"%x")
  years <- c((year-11):year)
  
  ## Get DV input file and template file
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("PM25",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.2012 <- get.naa.info(par=88101,psid=22)
  naa.2006 <- get.naa.info(par=88101,psid=16)
  naa.1997 <- get.naa.info(par=88101,psid=18)
  naa.states <- union(get.naa.states(naa.2012),union(get.naa.states(naa.2006),get.naa.states(naa.1997)))
  naa.states <- naa.states[order(naa.states)]
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=88101,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_1997",
    "naa_name_2006","naa_name_2012")]
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
  table0 <- data.frame(parameter="88101",site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,collection_frequency=cfr,
    season_begin_date=sbd,season_end_date=sed,monitor_types=types,monitor_networks=nets,
    measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  writeData(wb,sheet=15,x=table0,startCol=1,startRow=3,colNames=FALSE,rowNames=FALSE,na.string="")
  clear.rows <- c((nrow(table0)+3):max(wb$worksheets[[15]]$sheet_data$rows))
  removeRowHeights(wb,sheet=15,rows=clear.rows)
  deleteData(wb,sheet=15,rows=clear.rows,cols=c(1:ncol(table0)),gridExpand=TRUE)
  for (i in 1:ncol(table0)) {
    addStyle(wb,sheet=15,style=createStyle(border=c("top","bottom","left","right"),borderStyle="none"),
      rows=clear.rows,cols=i)
  }
  addStyle(wb,sheet=15,style=createStyle(border="top",borderColour="black",borderStyle="thin"),
    rows=min(clear.rows),cols=c(1:ncol(table0)))
  
  ## Pull annual and daily DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_PM25_ANNUAL_DVS
    WHERE dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '88101'
      AND pollutant_standard_id = 27
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year",sep=""),dbname="aqsprod")
  write.csv(t,file=paste("DVs2xlsx/",year,"/PM25anndvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dva <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv=as.numeric(t$design_value),valid=t$dv_validity_ind,
    mean.yr1=as.numeric(t$year_2_arith_mean),mean.yr2=as.numeric(t$year_1_arith_mean),
    mean.yr3=as.numeric(t$dv_year_arith_mean),qtrs.yr1=as.integer(t$year_2_complete_quarters),
    qtrs.yr2=as.integer(t$year_1_complete_quarters),qtrs.yr3=as.integer(t$dv_year_complete_quarters),
    mean.yr1.q1=as.numeric(t$yr2_q1_arith_mean),mean.yr1.q2=as.numeric(t$yr2_q2_arith_mean),
    mean.yr1.q3=as.numeric(t$yr2_q3_arith_mean),mean.yr1.q4=as.numeric(t$yr2_q4_arith_mean),
    mean.yr2.q1=as.numeric(t$yr1_q1_arith_mean),mean.yr2.q2=as.numeric(t$yr1_q2_arith_mean),
    mean.yr2.q3=as.numeric(t$yr1_q3_arith_mean),mean.yr2.q4=as.numeric(t$yr1_q4_arith_mean),
    mean.yr3.q1=as.numeric(t$dv_yr_q1_arith_mean),mean.yr3.q2=as.numeric(t$dv_yr_q2_arith_mean),
    mean.yr3.q3=as.numeric(t$dv_yr_q3_arith_mean),mean.yr3.q4=as.numeric(t$dv_yr_q4_arith_mean),
    obs.yr1.q1=as.integer(t$yr2_q1_creditable_cnt),obs.yr1.q2=as.integer(t$yr2_q2_creditable_cnt),
    obs.yr1.q3=as.integer(t$yr2_q3_creditable_cnt),obs.yr1.q4=as.integer(t$yr2_q4_creditable_cnt),
    obs.yr2.q1=as.integer(t$yr1_q1_creditable_cnt),obs.yr2.q2=as.integer(t$yr1_q2_creditable_cnt),
    obs.yr2.q3=as.integer(t$yr1_q3_creditable_cnt),obs.yr2.q4=as.integer(t$yr1_q4_creditable_cnt),
    obs.yr3.q1=as.integer(t$dv_yr_q1_creditable_cnt),obs.yr3.q2=as.integer(t$dv_yr_q2_creditable_cnt),
    obs.yr3.q3=as.integer(t$dv_yr_q3_creditable_cnt),obs.yr3.q4=as.integer(t$dv_yr_q4_creditable_cnt),
    pct.yr1.q1=pmin(round(100*as.integer(t$yr2_q1_creditable_cnt)/pmax(as.integer(t$yr2_q1_scheduled_cnt),1)),100),
    pct.yr1.q2=pmin(round(100*as.integer(t$yr2_q2_creditable_cnt)/pmax(as.integer(t$yr2_q2_scheduled_cnt),1)),100),
    pct.yr1.q3=pmin(round(100*as.integer(t$yr2_q3_creditable_cnt)/pmax(as.integer(t$yr2_q3_scheduled_cnt),1)),100),
    pct.yr1.q4=pmin(round(100*as.integer(t$yr2_q4_creditable_cnt)/pmax(as.integer(t$yr2_q4_scheduled_cnt),1)),100),
    pct.yr2.q1=pmin(round(100*as.integer(t$yr1_q1_creditable_cnt)/pmax(as.integer(t$yr1_q1_scheduled_cnt),1)),100),
    pct.yr2.q2=pmin(round(100*as.integer(t$yr1_q2_creditable_cnt)/pmax(as.integer(t$yr1_q2_scheduled_cnt),1)),100),
    pct.yr2.q3=pmin(round(100*as.integer(t$yr1_q3_creditable_cnt)/pmax(as.integer(t$yr1_q3_scheduled_cnt),1)),100),
    pct.yr2.q4=pmin(round(100*as.integer(t$yr1_q4_creditable_cnt)/pmax(as.integer(t$yr1_q4_scheduled_cnt),1)),100),
    pct.yr3.q1=pmin(round(100*as.integer(t$dv_yr_q1_creditable_cnt)/pmax(as.integer(t$dv_yr_q1_scheduled_cnt),1)),100),
    pct.yr3.q2=pmin(round(100*as.integer(t$dv_yr_q2_creditable_cnt)/pmax(as.integer(t$dv_yr_q2_scheduled_cnt),1)),100),
    pct.yr3.q3=pmin(round(100*as.integer(t$dv_yr_q3_creditable_cnt)/pmax(as.integer(t$dv_yr_q3_scheduled_cnt),1)),100),
    pct.yr3.q4=pmin(round(100*as.integer(t$dv_yr_q4_creditable_cnt)/pmax(as.integer(t$dv_yr_q4_scheduled_cnt),1)),100),
    obs.q1=as.integer(t$q1_3yr_creditable_cnt),obs.q2=as.integer(t$q2_3yr_creditable_cnt),
    obs.q3=as.integer(t$q3_3yr_creditable_cnt),obs.q4=as.integer(t$q4_3yr_creditable_cnt),
    max.q1=as.integer(t$q1_3yr_maximum),max.q2=as.integer(t$q2_3yr_maximum),
    max.q3=as.integer(t$q3_3yr_maximum),max.q4=as.integer(t$q4_3yr_maximum),
    min.q1=as.integer(t$q1_3yr_minimum),min.q2=as.integer(t$q2_3yr_minimum),
    min.q3=as.integer(t$q3_3yr_minimum),min.q4=as.integer(t$q4_3yr_minimum))
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_PM25_24HR_DVS
    WHERE dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '88101'
      AND pollutant_standard_id = 26
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year",sep=""),dbname="aqsprod")
  write.csv(t,file=paste("DVs2xlsx/",year,"/PM25_24hdvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvd <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv=as.numeric(t$daily_design_value),valid=t$dv_validity_ind,
    p98.yr1=as.numeric(t$year_2_98th_percentile),p98.yr2=as.numeric(t$year_1_98th_percentile),
    p98.yr3=as.numeric(t$dv_year_98th_percentile),qtrs.yr1=NA,qtrs.yr2=NA,qtrs.yr3=NA,
    obs.yr1=as.integer(t$year_2_creditable_cnt),obs.yr2=as.integer(t$year_1_creditable_cnt),
    obs.yr3=as.integer(t$dv_year_creditable_cnt),obs.yr1.q1=as.integer(t$yr2_q1_creditable_cnt),
    obs.yr1.q2=as.integer(t$yr2_q2_creditable_cnt),obs.yr1.q3=as.integer(t$yr2_q3_creditable_cnt),
    obs.yr1.q4=as.integer(t$yr2_q4_creditable_cnt),obs.yr2.q1=as.integer(t$yr1_q1_creditable_cnt),
    obs.yr2.q2=as.integer(t$yr1_q2_creditable_cnt),obs.yr2.q3=as.integer(t$yr1_q3_creditable_cnt),
    obs.yr2.q4=as.integer(t$yr1_q4_creditable_cnt),obs.yr3.q1=as.integer(t$dv_yr_q1_creditable_cnt),
    obs.yr3.q2=as.integer(t$dv_yr_q2_creditable_cnt),obs.yr3.q3=as.integer(t$dv_yr_q3_creditable_cnt),
    obs.yr3.q4=as.integer(t$dv_yr_q4_creditable_cnt),
    pct.yr1.q1=pmin(round(100*as.integer(t$yr2_q1_creditable_cnt)/pmax(as.integer(t$yr2_q1_scheduled_samples),1)),100),
    pct.yr1.q2=pmin(round(100*as.integer(t$yr2_q2_creditable_cnt)/pmax(as.integer(t$yr2_q2_scheduled_samples),1)),100),
    pct.yr1.q3=pmin(round(100*as.integer(t$yr2_q3_creditable_cnt)/pmax(as.integer(t$yr2_q3_scheduled_samples),1)),100),
    pct.yr1.q4=pmin(round(100*as.integer(t$yr2_q4_creditable_cnt)/pmax(as.integer(t$yr2_q4_scheduled_samples),1)),100),
    pct.yr2.q1=pmin(round(100*as.integer(t$yr1_q1_creditable_cnt)/pmax(as.integer(t$yr1_q1_scheduled_samples),1)),100),
    pct.yr2.q2=pmin(round(100*as.integer(t$yr1_q2_creditable_cnt)/pmax(as.integer(t$yr1_q2_scheduled_samples),1)),100),
    pct.yr2.q3=pmin(round(100*as.integer(t$yr1_q3_creditable_cnt)/pmax(as.integer(t$yr1_q3_scheduled_samples),1)),100),
    pct.yr2.q4=pmin(round(100*as.integer(t$yr1_q4_creditable_cnt)/pmax(as.integer(t$yr1_q4_scheduled_samples),1)),100),
    pct.yr3.q1=pmin(round(100*as.integer(t$dv_yr_q1_creditable_cnt)/pmax(as.integer(t$dv_yr_q1_scheduled_samples),1)),100),
    pct.yr3.q2=pmin(round(100*as.integer(t$dv_yr_q2_creditable_cnt)/pmax(as.integer(t$dv_yr_q2_scheduled_samples),1)),100),
    pct.yr3.q3=pmin(round(100*as.integer(t$dv_yr_q3_creditable_cnt)/pmax(as.integer(t$dv_yr_q3_scheduled_samples),1)),100),
    pct.yr3.q4=pmin(round(100*as.integer(t$dv_yr_q4_creditable_cnt)/pmax(as.integer(t$dv_yr_q4_scheduled_samples),1)),100),
    obs.q1=as.integer(t$yr2_q1_creditable_cnt)+as.integer(t$yr1_q1_creditable_cnt)+as.integer(t$dv_yr_q1_creditable_cnt),
    obs.q2=as.integer(t$yr2_q2_creditable_cnt)+as.integer(t$yr1_q2_creditable_cnt)+as.integer(t$dv_yr_q2_creditable_cnt),
    obs.q3=as.integer(t$yr2_q3_creditable_cnt)+as.integer(t$yr1_q3_creditable_cnt)+as.integer(t$dv_yr_q3_creditable_cnt),
    obs.q4=as.integer(t$yr2_q4_creditable_cnt)+as.integer(t$yr1_q4_creditable_cnt)+as.integer(t$dv_yr_q4_creditable_cnt),
    max.q1=as.numeric(t$q1_3yr_max),max.q2=as.numeric(t$q2_3yr_max),
    max.q3=as.numeric(t$q3_3yr_max),max.q4=as.numeric(t$q4_3yr_max))
  dvd$qtrs.yr1 <- apply(dvd[,c("pct.yr1.q1","pct.yr1.q2","pct.yr1.q3","pct.yr1.q4")],1,
    function(x) ifelse(all(is.na(x)),NA,sum(x >= 75,na.rm=TRUE)))
  dvd$qtrs.yr2 <- apply(dvd[,c("pct.yr2.q1","pct.yr2.q2","pct.yr2.q3","pct.yr2.q4")],1,
    function(x) ifelse(all(is.na(x)),NA,sum(x >= 75,na.rm=TRUE)))
  dvd$qtrs.yr3 <- apply(dvd[,c("pct.yr3.q1","pct.yr3.q2","pct.yr3.q3","pct.yr3.q4")],1,
    function(x) ifelse(all(is.na(x)),NA,sum(x >= 75,na.rm=TRUE)))
  sites <- subset(table0,!duplicated(site),c("site","epa_region","state_name","county_name","cbsa_name","csa_name",
    "naa_name_1997","naa_name_2006","naa_name_2012","site_name","address","latitude","longitude"))
  dvs.ann <- merge(sites,dva); dvs.24h <- merge(sites,dvd);
  
  ## Table 1a. NAA Status Annual 2012
  t <- subset(dvs.ann,naa_name_2012 != " " & dv_year == year & valid == "Y",c("naa_name_2012","dv"))
  t <- subset(t[order(t$naa_name_2012,t$dv,decreasing=TRUE),],!duplicated(naa_name_2012))
  table1a <- merge(naa.2012,t,by.x="naa_name",by.y="naa_name_2012",all=TRUE)
  table1a$met_naaqs <- sapply(table1a$dv,function(x) 
    ifelse(is.na(x),"Incomplete",ifelse(x > 12,"No","Yes")))
  table1a <- table1a[,c("naa_name","epa_regions","status","dv","met_naaqs","cdd_date","redesignation_date")]
  writeData(wb,sheet=1,x=table1a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 1b. NAA Status 24-hour 2006
  t <- subset(dvs.24h,naa_name_2006 != " " & dv_year == year & valid == "Y",c("naa_name_2006","dv"))
  t <- subset(t[order(t$naa_name_2006,t$dv,decreasing=TRUE),],!duplicated(naa_name_2006))
  table1b <- merge(naa.2006,t,by.x="naa_name",by.y="naa_name_2006",all=TRUE)
  table1b$met_naaqs <- sapply(table1b$dv,function(x) 
    ifelse(is.na(x),"Incomplete",ifelse(x > 35,"No","Yes")))
  table1b <- table1b[,c("naa_name","epa_regions","status","dv","met_naaqs","cdd_date","redesignation_date")]
  writeData(wb,sheet=2,x=table1b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  
  ## Table 1c. NAA Status Annual 1997
  t <- subset(dvs.ann,naa_name_1997 != " " & dv_year == year & valid == "Y",c("naa_name_1997","dv"))
  t <- subset(t[order(t$naa_name_1997,t$dv,decreasing=TRUE),],!duplicated(naa_name_1997))
  table1c <- merge(naa.1997,t,by.x="naa_name",by.y="naa_name_1997",all=TRUE)
  table1c$met_naaqs <- sapply(table1c$dv,function(x) 
    ifelse(is.na(x),"Incomplete",ifelse(x > 15,"No","Yes")))
  table1c <- table1c[,c("naa_name","epa_regions","status","dv","met_naaqs","cdd_date","redesignation_date")]
  writeData(wb,sheet=3,x=table1c,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 2a. Other Violators Annual
  t <- subset(dvs.ann,naa_name_2012 == " " & dv_year == year & dv > 12 & valid == "Y")
  table2a <- t[,c("state_name","county_name","epa_region","site","dv","cbsa_name")]
  if (nrow(table2a) == 0) {
    table2a <- data.frame(x=paste("There were no sites violating the 2012 Annual PM2.5 NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  writeData(wb,sheet=4,x=table2a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=4,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  remove.extra.rows(wb,sheet=4,row.hts=c(15,34,15,47,15,63,15))
  
  ## Table 2b. Other Violators 24-hour
  t <- subset(dvs.24h,naa_name_2006 == " " & dv_year == year & dv > 35 & valid == "Y")
  table2b <- t[,c("state_name","county_name","epa_region","site","dv","cbsa_name")]
  if (nrow(table2b) == 0) {
    table2b <- data.frame(x=paste("There were no sites violating the 2006 24-hour PM2.5 NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  writeData(wb,sheet=5,x=table2b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,34,15,47,15,63,15))
  
  ## Table 3a. NAA Trends Annual 2012
  temp <- subset(dvs.ann,naa_name_2012 != " " & valid == "Y")
  table3a <- data.frame(naa_name=naa.2012$naa_name,epa_region=naa.2012$epa_regions)
  for (y in years[3:length(years)]) {
    table3a[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_2012","dv"))
    for (i in 1:nrow(naa.2012)) {
      v <- subset(t,naa_name_2012 == naa.2012$naa_name[i])
      if (nrow(v) == 0) { next }
      table3a[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=6,x=table3a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=6,years); set.footnote.dates(wb,sheet=6);
  
  ## Table 3b. NAA Trends 24-hour 2006
  temp <- subset(dvs.24h,naa_name_2006 != " " & valid == "Y")
  table3b <- data.frame(naa_name=naa.2006$naa_name,epa_regions=naa.2006$epa_regions)
  for (y in years[3:length(years)]) {
    table3b[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_2006","dv"))
    for (i in 1:nrow(naa.2006)) {
      v <- subset(t,naa_name_2006 == naa.2006$naa_name[i])
      if (nrow(v) == 0) { next }
      table3b[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=7,x=table3b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=7,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=7,years); set.footnote.dates(wb,sheet=7);
  
  ## Table 3c. NAA Trends Annual 1997
  temp <- subset(dvs.ann,naa_name_1997 != " " & valid == "Y")
  table3c <- data.frame(naa_name=naa.1997$naa_name,epa_regions=naa.1997$epa_regions)
  for (y in years[3:length(years)]) {
    table3c[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name_1997","dv"))
    for (i in 1:nrow(naa.1997)) {
      v <- subset(t,naa_name_1997 == naa.1997$naa_name[i])
      if (nrow(v) == 0) { next }
      table3c[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv)
    }
  }
  writeData(wb,sheet=8,x=table3c,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=8,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=8,years); set.footnote.dates(wb,sheet=8);
  
  ## Table 4a. County Status Annual
  t <- subset(dvs.ann,dv_year == year & valid == "Y",
    c("site","state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,dv=t$dv,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  writeData(wb,sheet=9,x=table4a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=9,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=9,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=9,years); set.footnote.dates(wb,sheet=9);
  remove.extra.rows(wb,sheet=9,row.hts=c(15,34,15,47,15,63,15))
  
  ## Table 4b. County Status 24-hour
  t <- subset(dvs.24h,dv_year == year & valid == "Y",
    c("site","state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,dv=t$dv,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  writeData(wb,sheet=10,x=table4b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=10,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=10,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=10,years); set.footnote.dates(wb,sheet=10);
  remove.extra.rows(wb,sheet=10,row.hts=c(15,34,15,47,15,63,15))
  
  ## Table 5a. Site Status Annual
  t <- subset(dvs.ann,dv_year == year & obs.q1 + obs.q2 + obs.q3 + obs.q4 > 0)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv,NA),t$dv,t$valid)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2012","naa_name_1997",
    "epa_region","site","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    paste(rep(c("mean","qtrs"),each=3),rep(paste("yr",1:3,sep=""),times=2),sep="."),
    paste(rep(c("mean","obs","pct"),each=12),rep(paste("yr",1:3,sep=""),each=4,times=3),
      rep(paste("q",1:4,sep=""),times=9),sep="."),
    paste(rep(c("obs","max","min"),each=4),rep(paste("q",1:4,sep=""),times=3),sep="."))]
  writeData(wb,sheet=11,x=table5a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=11,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=11,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=11,years); set.footnote.dates(wb,sheet=11);
  remove.extra.rows(wb,sheet=11,row.hts=c(15,34,15,47,15,79,15))
  
  ## Table 5b. Site Status 24-hour
  t <- subset(dvs.24h,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv,NA),t$dv,t$valid)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2006",
    "epa_region","site","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    paste(rep(c("p98","obs","qtrs"),each=3),rep(paste("yr",1:3,sep=""),times=3),sep="."),
    paste(rep(c("obs","pct"),each=12),rep(paste("yr",1:3,sep=""),each=4,times=2),
      rep(paste("q",1:4,sep=""),times=6),sep="."),
    paste(rep(c("obs","max"),each=4),rep(paste("q",1:4,sep=""),times=2),sep="."))]
  writeData(wb,sheet=12,x=table5b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=12,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=12,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=12,years); set.footnote.dates(wb,sheet=12);
  remove.extra.rows(wb,sheet=12,row.hts=c(15,34,15,47,15,79,15))
  
  ## Table 6a. Site Trends Annual
  t <- dcast(subset(dvs.ann,valid == "Y"),site ~ dv_year,value.var="dv")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs.ann,!duplicated(site)),t,by="site")
  table6a <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2012",
    "naa_name_1997","epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=13,x=table6a,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=13,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=13,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=13,years); set.footnote.dates(wb,sheet=13);
  remove.extra.rows(wb,sheet=13,row.hts=c(15,34,15,47,15,79,15))
  
  ## Table 6b. Site Trends 24-hour
  t <- dcast(subset(dvs.24h,valid == "Y"),site ~ dv_year,value.var="dv")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs.24h,!duplicated(site)),t,by="site")
  table6b <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name_2006",
    "epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=14,x=table6b,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=14,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=14,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=14,years); set.footnote.dates(wb,sheet=14);
  remove.extra.rows(wb,sheet=14,row.hts=c(15,34,15,47,15,79,15))
  
  ## Write DV tables to Excel File
  fix.scripts(wb)
  file.xlsx <- paste("PM25_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}