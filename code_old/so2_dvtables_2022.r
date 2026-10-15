## Format SO2 DV spreadsheets based on template file
so2.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("SO2",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42401,psid=19)
  naa.info$type <- sapply(naa.info$designation_date,function(x) 
   ifelse(as.numeric(substr(gsub(" ",Sys.Date(),x),1,4)) %in% c(2016:2018),"Modeled","Monitored"))
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42401,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_2010")]
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
  table0 <- data.frame(parameter="42401",site=substr(ids,1,9),poc=substr(ids,10,10),
    base_info,monitor_begin_date=mbd,monitor_end_date=med,last_sample_date=lsd,
    primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,nonreg_end_date=ned,
    nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,collection_frequency=cfr,
    season_begin_date=sbd,season_end_date=sed,monitor_types=types,monitor_networks=nets,
    measurement_scale=msc,monitor_objective=obj,row.names=NULL)
  df2xls(df=table0,sheet=dv.sheets[[7]],sr=3,sc=1)
  
  ## Pull 1-hr DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_SO2_DVS
    WHERE dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '42401'
      AND pollutant_standard_id = 19
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year"),dbname="aqsprod")
  write.csv(t,file=paste("DVs2xlsx/",year,"/SO2dvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.so2 <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=t$dv_year,dv=as.numeric(t$design_value),valid=t$dv_validity_indicator,
    ind.yr1=t$year_2_complete_ind,ind.yr2=t$year_1_complete_ind,ind.yr3=t$year_0_complete_ind,
    qtr.yr1=t$year_2_complete_quarters,qtr.yr2=t$year_1_complete_quarters,
    qtr.yr3=t$year_0_complete_quarters,p99.yr1=as.numeric(t$year_2_99th_percentile),
    p99.yr2=as.numeric(t$year_1_99th_percentile),p99.yr3=as.numeric(t$year_0_99th_percentile))
  sites <- subset(table0,!duplicated(site),c("site","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude"))
  dvs <- merge(sites,dvs.so2)
  
  ## Table 1. NAA Status
  t <- subset(dvs,naa_name != " " & dv_year == year & valid == "Y",c("naa_name","dv"))
  t <- subset(t[order(t$naa_name,t$dv,decreasing=TRUE),],!duplicated(naa_name))
  table1c <- merge(naa.info,t,by="naa_name",all=TRUE)
  monitored <- subset(table1c,type == "Monitored",c("naa_name","epa_regions","status","dv"))
  monitored$met_naaqs <- sapply(monitored$dv,function(x) ifelse(is.na(x),"Incomplete",
    ifelse(x > 75,"No","Yes")))
  modeled <- subset(table1c,type == "Modeled",c("naa_name","epa_regions","status","dv"))
  modeled$met_naaqs <- sapply(modeled$dv,function(x) ifelse(is.na(x),"[4]",
    ifelse(x > 75,"No","[4]")))
  set.aqs.date(dv.sheets[[1]],today,row=2,col=1)
  set.last.update(dv.sheets[[1]],row=2,col=2)
  set.year.values(dv.sheets[[1]],years,row=4)
  df2xls(df=monitored,sheet=dv.sheets[[1]],sr=6,sc=1)
  df2xls(df=modeled,sheet=dv.sheets[[1]],sr=44,sc=1)
  set.footnote.date(dv.sheets[[1]],today,row=58,col=1)
  set.footnote.date(dv.sheets[[1]],today,row=62,col=1)
  
  ## Table 2. Other Violators
  t <- subset(dvs,naa_name == " " & dv_year == year & dv > 75 & valid == "Y")
  table2 <- t[,c("state_name","county_name","epa_region","site","dv","cbsa_name")]
  if (nrow(table2) == 0) {
    table2 <- data.frame(x=paste("There were no sites violating the 1-hour SO2 NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  set.aqs.date(dv.sheets[[2]],today,row=2,col=1)
  set.last.update(dv.sheets[[2]],row=2,col=3)
  set.year.values(dv.sheets[[2]],years,row=4)
  df2xls(df=table2,sheet=dv.sheets[[2]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[2]],today,row=14,col=1)
  
  ## Table 3. NAA Trends
  temp <- subset(dvs,naa_name != " " & valid == "Y")
  table3 <- data.frame(naa_name=naa.info$naa_name,epa_regions=naa.info$epa_regions)
  for (y in years[3:length(years)]) {
    table3[,paste("dv",(y-2),y,sep="_")] <- NA
    t <- subset(temp,dv_year == y,c("naa_name","dv"))
    for (i in 1:nrow(naa.info)) {
      v <- subset(t,naa_name == naa.info$naa_name[i])
      if (nrow(v) == 0) { next }
      table3[i,paste("dv",(y-2),y,sep="_")] <- max(v$dv)
    }
  }
  set.aqs.date(dv.sheets[[3]],today,row=2,col=1)
  set.last.update(dv.sheets[[3]],row=2,col=2)
  set.year.values(dv.sheets[[3]],years,row=4)
  df2xls(df=table3,sheet=dv.sheets[[3]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[3]],today,row=58,col=1)
  
  ## Table 4. County Status
  t <- subset(dvs,dv_year == year & valid == "Y",
    c("site","state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4 <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv,cbsa_name=t$cbsa_name)
  table4 <- table4[order(table4$state_name,table4$county_name),]
  set.aqs.date(dv.sheets[[4]],today,row=2,col=1)
  set.last.update(dv.sheets[[4]],row=2,col=3)
  set.year.values(dv.sheets[[4]],years,row=4)
  df2xls(df=table4,sheet=dv.sheets[[4]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[4]],today,row=289,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[4]],row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5. Site Status
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv,t$valid)
  table5 <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv","ind.yr1",
    "ind.yr2","ind.yr3","qtr.yr1","qtr.yr2","qtr.yr3","p99.yr1","p99.yr2","p99.yr3")]
  set.aqs.date(dv.sheets[[5]],today,row=2,col=1)
  set.last.update(dv.sheets[[5]],row=2,col=3)
  set.year.values(dv.sheets[[5]],years,row=4)
  df2xls(df=table5,sheet=dv.sheets[[5]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[5]],today,row=459,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[5]],row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 6. Site Trends
  t <- dcast(subset(dvs,valid == "Y"),site ~ dv_year,value.var="dv")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6 <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  set.aqs.date(dv.sheets[[6]],today,row=2,col=1)
  set.last.update(dv.sheets[[6]],row=2,col=3)
  set.year.values(dv.sheets[[6]],years,row=4)
  df2xls(df=table6,sheet=dv.sheets[[6]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[6]],today,row=589,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[6]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Write DV tables to Excel File
  file.xlsx <- paste("SO2_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(dv.wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""))
}