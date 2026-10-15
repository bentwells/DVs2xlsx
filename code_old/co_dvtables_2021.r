## Format CO DV spreadsheets based on template file
co.dvtables <- function(year,type="DRAFT") {
  
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
  source("C:/Users/bwells01/Documents/R/xlsx_dv_funs.r")
  require(plyr); require(reshape2); require(xlsx);
  step1.date <- format(Sys.Date(),"%d%b%y"); years <- c((year-10):year);
  dv.begin.date <- as.Date(paste(year-1,"01-01",sep="-"))
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  template.file <- templates[intersect(grep("CO",templates),grep(type,templates))]
  dv.wb <- loadWorkbook(paste("DVs2xlsx/templates",template.file,sep="/"))
  dv.sheets <- getSheets(dv.wb)

  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42101,psid=4)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42101,yr1=years[1],yr2=year,all=TRUE)
  ids <- monitors$id[!duplicated(monitors$id)]
  agency <- monitors$reporting_agency[!duplicated(monitors$id)]
  base_info <- monitors[!duplicated(monitors$id),c("epa_region","state_name","county_name",
    "cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_1971")]
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
  table0 <- data.frame(parameter="42101",site=substr(ids,1,9),poc=substr(ids,10,10),row.names=NULL)
  if (type == "DRAFT") {
    v_area <- ifelse(as.Date(mbd) >= as.Date(paste(year,"01-01",sep="-")) &
      base_info$state_name %in% naa.states,"1"," ")
    v_nonreg <- ifelse(nbd != " " & nrc == " " & as.Date(lsd) >= dv.begin.date,"1"," ")
    v_nonref <- ifelse(non_ref != " " & nrc != "Y" & check.dates == TRUE &
      as.Date(lsd) >= dv.begin.date,"1"," ")
    v_closure <- ifelse(med == " " & nrc != "Y" & as.Date(lsd) >= dv.begin.date &
      as.Date(lsd) <= as.Date(paste(year,"10-01",sep="-")),"1"," ")
    v_spm <- ifelse(grepl("SPM",types) & nrc != "Y" & as.Date(lsd) >= dv.begin.date,
      ifelse(med == " ",ifelse(as.Date(lsd) - as.Date(mbd) <= 730,"1"," "),
      ifelse(as.Date(gsub(" ",as.character(Sys.Date()),med)) - as.Date(mbd) <= 730,"1"," "))," ")
    v_combo <- ifelse(combo_site != " " & as.Date(gsub(" ","1999-01-01",combo_date)) >= dv.begin.date,"1"," ")
    v_count <- (v_area != " ") + (v_nonreg != " ") + (v_nonref != " ") + (v_closure != " ") +
      (v_spm != " ") + (v_combo != " ")
    table0 <- cbind(table0,data.frame(v_count,v_area,v_nonreg,v_nonref,v_closure,v_spm,v_combo,row.names=NULL))
  }
  table0 <- cbind(table0,data.frame(base_info,monitor_begin_date=mbd,monitor_end_date=med,
    last_sample_date=lsd,primary_begin_date=pbd,primary_end_date=ped,nonreg_begin_date=nbd,
    nonreg_end_date=ned,nonreg_concur=nrc,frm_fem,non_ref,combo_site,combo_date,agency,
    collection_frequency=cfr,season_begin_date=sbd,season_end_date=sed,monitor_types=types,
    monitor_networks=nets,measurement_scale=msc,monitor_objective=obj,row.names=NULL))
  S <- ifelse(type == "DRAFT",1,length(dv.sheets))
  df2xls(df=table0,sheet=dv.sheets[[S]],sr=3,sc=1)
  
  ## Pull CO DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_CO_DVS
    WHERE parameter_code = '42101'
      AND dv_year >=",years[2],
     "AND dv_year <=",year,
     "AND edt_id IN (0,5)
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, poc, dv_year",sep=""))
  write.csv(t,file=paste("DVs2xlsx/",year,"/COdvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.co <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    poc=t$poc,dv_year=t$dv_year,
    dv_1hr=t$co_1hr_2nd_max_value,dt_1hr=as.character(t$co_8hr_2nd_max_date_time),
    dv_8hr=t$co_8hr_2nd_max_value,dt_8hr=as.character(t$co_8hr_2nd_max_date_time))   
  sites <- table0[,c("site","poc","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude")]
  dvs <- merge(sites,dvs.co,by=c("site","poc"))
  dvs$valid <- TRUE
  spm.check <- paste(table0$site,table0$poc,sep="")[which(grepl("SPM",table0$monitor_types) & 
    (as.Date(gsub(" ",Sys.Date(),table0$monitor_end_date)) - as.Date(table0$monitor_begin_date) <= 730 |
     as.Date(table0$last_sample_date) - as.Date(table0$monitor_begin_date) <= 730))]
  dvs$valid[which(paste(dvs$site,dvs$poc,sep="") %in% spm.check)] <- FALSE
  
  ## Table 1a. NAA Status 8hr
  t <- subset(dvs,naa_name != " " & dv_year == year & valid,c("naa_name","dv_8hr"))
  t <- subset(t[order(t$naa_name,t$dv_8hr,decreasing=TRUE),],!duplicated(naa_name))
  table1a <- merge(naa.info,t,by="naa_name",all=TRUE)[,c("naa_name","epa_regions","status","dv_8hr")]
  table1a$met_naaqs <- sapply(table1a$dv_8hr,function(x) ifelse(is.na(x),"No Data",
    ifelse(x > 9,"No","Yes")))
  S <- ifelse(type == "DRAFT",2,1)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table1a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=84,col=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=88,col=1)
  
  ## Table 1b. NAA Status 1hr
  t <- subset(dvs,naa_name != " " & dv_year == year & valid,c("naa_name","dv_1hr"))
  t <- subset(t[order(t$naa_name,t$dv_1hr,decreasing=TRUE),],!duplicated(naa_name))
  table1b <- merge(naa.info,t,by="naa_name",all=TRUE)[,c("naa_name","epa_regions","status","dv_1hr")]
  table1b$met_naaqs <- sapply(table1b$dv_1hr,function(x) ifelse(is.na(x),"No Data",
    ifelse(x > 35,"No","Yes")))
  S <- ifelse(type == "DRAFT",3,2)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table1b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=84,col=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=88,col=1)
  
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
  S <- ifelse(type == "DRAFT",4,3)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table2a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
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
  S <- ifelse(type == "DRAFT",5,4)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table2b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
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
  S <- ifelse(type == "DRAFT",6,5)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table3a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=86,col=1)
  
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
  S <- ifelse(type == "DRAFT",7,6)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table3b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=86,col=1)
  
  ## Table 4a. County Status 8hr
  t <- subset(dvs,dv_year == year & valid,c("site","poc","state_name","county_name",
    "epa_region","dv_8hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_8hr,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,poc=t$poc,dv=t$dv_8hr,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  S <- ifelse(type == "DRAFT",8,7)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table4a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=189,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,63,15,63,15,95,15))
  
  ## Table 4b. County Status 1hr
  t <- subset(dvs,dv_year == year & valid,c("site","poc","state_name","county_name",
    "epa_region","dv_1hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_1hr,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),epa_region=t$epa_region,
    site=t$site,poc=t$poc,dv=t$dv_1hr,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  S <- ifelse(type == "DRAFT",9,8)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table4b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=189,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,63,15,63,15,95,15))
  
  ## Table 5a. Monitor Status 8hr
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid,dv,NA),t$dv_8hr,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid,NA,dv),t$dv_8hr,t$valid)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","poc","site_name","address","latitude","longitude","valid_dv","invalid_dv","dt_8hr")]
  S <- ifelse(type == "DRAFT",10,9)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table5a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=309,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,47,15,63,15,79,15))
  
  ## Table 5b. Monitor Status 1hr
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid,dv,NA),t$dv_1hr,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid,NA,dv),t$dv_1hr,t$valid)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","poc","site_name","address","latitude","longitude","valid_dv","invalid_dv","dt_1hr")]
  S <- ifelse(type == "DRAFT",11,10)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table5b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=309,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,47,15,63,15,79,15))
  
  ## Table 6a. Monitor Trends 8hr
  t <- dcast(subset(dvs,valid),site + poc ~ dv_year,value.var="dv_8hr")
  colnames(t)[3:12] <- paste("dv",years[1:10],years[2:11],sep="_")
  table6a <- merge(sites,t,by=c("site","poc"))[,c("state_name","county_name","cbsa_name",
    "csa_name","naa_name","epa_region","site","poc","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[2:11],sep="_"))]
  S <- ifelse(type == "DRAFT",12,11)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table6a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,47,15,63,15,79,15))
  
  ## Table 6b. Monitor Trends 1hr
  t <- dcast(subset(dvs,valid),site + poc ~ dv_year,value.var="dv_1hr")
  colnames(t)[3:12] <- paste("dv",years[1:10],years[2:11],sep="_")
  table6b <- merge(sites,t,by=c("site","poc"))[,c("state_name","county_name","cbsa_name",
    "csa_name","naa_name","epa_region","site","poc","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[2:11],sep="_"))]
  S <- ifelse(type == "DRAFT",13,12)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4,co=TRUE)
  df2xls(df=table6b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,47,15,63,15,79,15))
  
  ## Write DV tables to Excel File
  file.xlsx <- paste("CO_DesignValues",year-1,year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(dv.wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""))
}