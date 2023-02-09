## Format NO2 DV spreadsheets based on template file
no2.dvtables <- function(year,type="DRAFT") {
  
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
  source("C:/Users/bwells01/Documents/R/xlsx_dv_funs.r")
  require(plyr); require(reshape2); require(xlsx);
  step1.date <- format(Sys.Date(),"%d%b%y"); years <- c((year-11):year);
  dv.begin.date <- as.Date(paste(year-2,"01-01",sep="-"))
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  template.file <- templates[intersect(grep("NO2",templates),grep(type,templates))]
  dv.wb <- loadWorkbook(paste("DVs2xlsx/templates",template.file,sep="/"))
  dv.sheets <- getSheets(dv.wb)
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42602,psid=8)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42602,yr1=years[1],yr2=year,all=TRUE)
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
  table0 <- data.frame(parameter="42602",site=substr(ids,1,9),poc=substr(ids,10,10),row.names=NULL)
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
    dv_year=t$dv_year,dv_ann=round(t$design_value),pct_ann=t$observation_percent,
    valid_ann=t$dv_validity_indicator)
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
    dv_year=t$dv_year,dv_1hr=t$design_value,valid_1hr=t$dv_validity_indicator,
    qtrs_yr1=t$year_2_complete_quarters,qtrs_yr2=t$year_1_complete_quarters,
    qtrs_yr3=t$year_0_complete_quarters,p98_yr1=t$year_2_98th_percentile,
    p98_yr2=t$year_1_98th_percentile,p98_yr3=t$year_0_98th_percentile)
  sites <- table0[!duplicated(table0$site),c("site","epa_region","state_name","county_name",
    "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude")]
  dvs <- merge(sites,merge(dvs.ann,dvs.1hr,by=c("site","dv_year"),all=TRUE),by="site")
  dvs$dv_1hr[which(is.na(dvs$p98_yr1) | is.na(dvs$p98_yr2) | is.na(dvs$p98_yr3))] <- NA
  dvs <- dvs[order(dvs$site,dvs$dv_year),]
  
  ## Table 1a: Nonattainment area status for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,naa_name != " " & dv_year == year & valid_ann == "Y",c("naa_name","dv_ann"))
  t <- subset(t[order(t$naa_name,t$dv_ann,decreasing=TRUE),],!duplicated(naa_name))
  table1a <- merge(naa.info[,c("naa_name","epa_regions","status")],t,by="naa_name",all=TRUE)
  table1a$meets_naaqs <- sapply(table1a$dv,function(x) ifelse(x > 53,"No","Yes"))
  S <- ifelse(type == "DRAFT",2,1)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table1a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=8,col=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=12,col=1)
  
  ## Table 2a: Additional monitors violating the 1971 Annual NO2 NAAQS
  table2a <- subset(dvs,naa_name != " " & dv_ann > 53 & valid_ann == "Y",
    c("state_name","county_name","epa_region","site","dv_ann","cbsa_name"))
  if (nrow(table2a) == 0) {
    table2a <- data.frame(x=paste("There were no sites violating the annual NO2 NAAQS in ",
      year,".",sep=""))
  }
  S <- ifelse(type == "DRAFT",3,2)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table2a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
  ## Table 2b: Monitors violating the 2010 1-hour NO2 NAAQS
  table2b <- subset(dvs,dv_1hr > 100 & valid_1hr == "Y", 
    c("state_name","county_name","epa_region","site","dv_1hr","cbsa_name"))
  if (nrow(table2b) == 0) {
    table2b <- data.frame(x=paste("There were no sites violating the 1-hour NO2 NAAQS in ",
      year-2,"-",year,".",sep=""))
  }
  S <- ifelse(type == "DRAFT",4,3)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table2b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
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
  S <- ifelse(type == "DRAFT",5,4)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table3a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
  ## Table 4a: County-level design values for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,dv_year == year & valid_ann == "Y",c("site",
    "state_name","county_name","epa_region","dv_ann","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4a <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$site,1,2),county_fips=substr(t$site,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv_ann,cbsa_name=t$cbsa_name)
  table4a <- table4a[order(table4a$state_name,table4a$county_name),]
  S <- ifelse(type == "DRAFT",6,5)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table4a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=259,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 4b: County-level design values for the 2010 1-hour NO2 NAAQS
  t <- subset(dvs,dv_year == year & valid_1hr == "Y",c("site",
    "state_name","county_name","epa_region","dv_1hr","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv_1hr,decreasing=TRUE),],!duplicated(fips))
  table4b <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$site,1,2),county_fips=substr(t$site,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv_1hr,cbsa_name=t$cbsa_name)
  table4b <- table4b[order(table4b$state_name,table4b$county_name),]
  S <- ifelse(type == "DRAFT",7,6)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table4b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=229,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5a: Site-level design values for the 1971 Annual NO2 NAAQS
  t <- subset(dvs,dv_year == year & pct_ann > 0)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv_ann,t$valid_ann)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv_ann,t$valid_ann)
  table5a <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv","pct_ann")]
  S <- ifelse(type == "DRAFT",8,7)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table5a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=489,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 5b: Site-level design values for the 2010 1-hour NO2 NAAQS
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv_1hr,t$valid_1hr)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv_1hr,t$valid_1hr)
  table5b <- t[,c("state_name","county_name","cbsa_name","csa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    "qtrs_yr1","qtrs_yr2","qtrs_yr3","p98_yr1","p98_yr2","p98_yr3")]
  S <- ifelse(type == "DRAFT",9,8)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table5b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6a: Trends in site-level design values for the 1971 Annual NO2 NAAQS
  t <- dcast(subset(dvs,valid_ann == "Y"),site ~ dv_year,value.var="dv_ann")
  colnames(t)[2:11] <- paste("dv",years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6a <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude",paste("dv",years[3:12],sep="_"))]
  S <- ifelse(type == "DRAFT",10,9)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table6a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=609,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6b: Trends in site-level design values for the 2010 1-hour NO2 NAAQS
  t <- dcast(subset(dvs,valid_1hr == "Y"),site ~ dv_year,value.var="dv_1hr")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6b <- vals[,c("state_name","county_name","cbsa_name","csa_name","epa_region","site",
    "site_name","address","latitude","longitude",paste("dv",years[1:10],years[3:12],sep="_"))]
  S <- ifelse(type == "DRAFT",11,10)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table6b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Save DV tables in .Rdata format and write to Excel file
  file.xlsx <- paste("NO2_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(dv.wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""))
}
