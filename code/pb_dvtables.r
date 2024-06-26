## Format Lead DV spreadsheets based on template file
pb.dvtables <- function(year=as.numeric(format(Sys.Date(),"%Y"))-1,
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
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  wb <- loadWorkbook(paste("DVs2xlsx/templates",templates[grep("Pb",templates)],sep="/"))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=14129,psid=2)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  t1 <- get.monitors(par=12128,yr1=years[1],yr2=year,all=TRUE)
  t2 <- get.monitors(par=14129,yr1=years[1],yr2=year,all=TRUE)
  t3 <- get.monitors(par=85129,yr1=years[1],yr2=year,all=TRUE)
  t1$pc <- "12128"; t2$pc <- "14129"; t3$pc <- "85129";
  monitors <- rbind(t1,t2,t3)
  monitors <- monitors[order(monitors$id,monitors$pc),]
  ids <- monitors$id[!duplicated(paste(monitors$id,monitors$pc))]
  pcs <- monitors$pc[!duplicated(paste(monitors$id,monitors$pc))]
  agency <- monitors$reporting_agency[!duplicated(paste(monitors$id,monitors$pc))]
  base_info <- monitors[!duplicated(paste(monitors$id,monitors$pc)),c("epa_region","state_name",
    "county_name","cbsa_name","csa_name","site_name","address","latitude","longitude","naa_name_2009")]
  base_info$latitude <- as.numeric(base_info$latitude)
  base_info$longitude <- as.numeric(base_info$longitude)
  colnames(base_info)[ncol(base_info)] <- "naa_name"
  mbd <- get.unique.dates(monitors$monitor_begin_date,paste(monitors$id,monitors$pc),first=TRUE)
  med <- get.unique.dates(monitors$monitor_end_date,paste(monitors$id,monitors$pc),first=FALSE)
  lsd <- get.unique.dates(monitors$last_sample_date,paste(monitors$id,monitors$pc),first=FALSE)
  pbd <- get.unique.dates(monitors$primary_begin_date,paste(monitors$id,monitors$pc),first=TRUE)
  ped <- get.unique.dates(monitors$primary_end_date,paste(monitors$id,monitors$pc),first=FALSE)
  sbd <- substr(get.unique.codes(monitors$season_begin_date,paste(monitors$id,monitors$pc),
    monitors$season_begin_date),6,10)
  sed <- substr(get.unique.codes(monitors$season_begin_date,paste(monitors$id,monitors$pc),
    monitors$season_end_date),6,10)
  nbd <- get.unique.dates(monitors$nonreg_begin_date,paste(monitors$id,monitors$pc),first=TRUE)
  ned <- get.unique.dates(monitors$nonreg_end_date,paste(monitors$id,monitors$pc),first=FALSE)
  nrc <- get.unique.codes(monitors$nonreg_begin_date,paste(monitors$id,monitors$pc),monitors$nonreg_concur)
  frc <- get.unique.codes(monitors$method_begin_date,paste(monitors$id,monitors$pc),monitors$frm_code)
  mtc <- get.unique.codes(monitors$method_begin_date,paste(monitors$id,monitors$pc),monitors$method_code)
  msc <- get.unique.codes(monitors$monitor_begin_date,paste(monitors$id,monitors$pc),monitors$measurement_scale)
  obj <- get.unique.codes(monitors$monitor_begin_date,paste(monitors$id,monitors$pc),monitors$monitor_objective)
  cfr <- get.unique.codes(monitors$monitor_begin_date,paste(monitors$id,monitors$pc),monitors$collection_frequency)
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
  types <- gsub(" ,","",tapply(monitors$monitor_type,list(paste(monitors$id,monitors$pc)),
    function(x) paste(unique(x[order(x)]),collapse=",")))
  if (any(types == " ")) { types[which(types == " ")] <- "OTHER" }
  monitors$network <- gsub("UNOFFICIAL ","",gsub("PROPOSED ","",monitors$network))
  nets <- tapply(monitors$network,list(paste(monitors$id,monitors$pc)),
    function(x) paste(unique(x[order(x)]),collapse=","))
  combos <- subset(monitors,combo_site != " ",c("id","combo_site","combo_date"))
  combos <- combos[which(!duplicated(combos)),]
  combo_site <- sapply(substr(ids,1,9),function(x) ifelse(x %in% substr(combos$id,1,9),
    combos$combo_site[match(x,substr(combos$id,1,9))],ifelse(x %in% combos$combo_site,
    substr(combos$id[match(x,combos$combo_site)],1,9)," ")))
  combo_date <- sapply(substr(ids,1,9),function(x) ifelse(x %in% substr(combos$id,1,9),
    combos$combo_date[match(x,substr(combos$id,1,9))],ifelse(x %in% combos$combo_site,
    combos$combo_date[match(x,combos$combo_site)]," ")))
  table0 <- data.frame(parameter=pcs,site=substr(ids,1,9),poc=as.numeric(substr(ids,10,11)),
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
  
  ## Pull Lead DVs from AQS, merge with site metadata
  t <- get.aqs.data(paste(
  "SELECT * FROM EUV_LEAD_DVS
    WHERE dv_year >=",years[3],"
      AND dv_year <=",year,"
      AND edt_id IN (0,5)
      AND parameter_code = '14129'
      AND pollutant_standard_id = 2
      AND state_code NOT IN ('80','CC')
    ORDER BY state_code, county_code, site_number, dv_year"))
  write.csv(t,file=paste("DVs2xlsx/",year,"/Pbdvs",year-9,"_",year,"_",
    format(Sys.Date(),"%Y%m%d"),".csv",sep=""),na="",row.names=FALSE)
  dvs.pb <- data.frame(site=paste(t$state_code,t$county_code,t$site_number,sep=""),
    dv_year=as.integer(t$dv_year),dv=as.numeric(t$design_value),valid=t$dv_validity_ind,
    valid_months=as.integer(t$total_valid_months),dv_month=paste(t$dv_max_month,t$dv_max_year,sep="-"),
    valid.yr1=as.integer(t$year_2_valid_months),valid.yr2=as.integer(t$year_1_valid_months),
    valid.yr3=as.integer(t$dv_year_valid_months),max.yr1=as.numeric(t$year_2_max_value),
    max.yr2=as.numeric(t$year_1_max_value),max.yr3=as.numeric(t$dv_year_max_value),
    month.yr1=as.integer(t$year_2_max_month),month.yr2=as.integer(t$year_1_max_month),
    month.yr3=as.integer(t$dv_year_max_month))
  sites <- subset(table0,!duplicated(site),c("site","epa_region","state_name","county_name",
   "cbsa_name","csa_name","naa_name","site_name","address","latitude","longitude"))
  dvs <- merge(sites,dvs.pb)
  
  ## Table 1. NAA Status
  naa.test <- subset(dvs,naa_name != " " & dv_year == year)
  t <- subset(dvs,naa_name != " " & dv_year == year & valid == "Y",c("naa_name","dv"))
  t <- subset(t[order(t$naa_name,t$dv,decreasing=TRUE),],!duplicated(naa_name))
  table1 <- merge(naa.info,t,by="naa_name",all=TRUE)[,c("naa_name","epa_regions","status","dv")]
  table1$met_naaqs <- mapply(function(naa,dv) ifelse(naa %in% unique(naa.test$naa_name),
    ifelse(is.na(dv),"Incomplete",ifelse(dv > 0.15,"No","Yes")),"No Data"),table1$naa_name,table1$dv)
  writeData(wb,sheet=1,x=table1,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=1,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=1,years); set.footnote.dates(wb,sheet=1);
  
  ## Table 2. Other Violators
  t <- subset(dvs,naa_name == " " & dv_year == year & dv > 0.15 & valid == "Y")
  table2 <- t[,c("state_name","county_name","epa_region","site","dv","cbsa_name")]
  if (nrow(table2) == 0) {
    table2 <- data.frame(x=paste("There were no sites violating the Lead NAAQS in ",
      (year-2),"-",year,".",sep=""))
  }
  writeData(wb,sheet=2,x=table2,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=2,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=2,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=2,years); set.footnote.dates(wb,sheet=2);
  
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
  writeData(wb,sheet=3,x=table3,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=3,x=t(paste(c("AQS Data Retrieval:","Last Updated:"),today)),
    startCol=1,startRow=2,colNames=FALSE,rowNames=FALSE,na.string="")
  set.year.values(wb,sheet=3,years); set.footnote.dates(wb,sheet=3);
  
  ## Table 4. County Status
  t <- subset(dvs,dv_year == year & valid == "Y",
    c("site","state_name","county_name","epa_region","dv","cbsa_name"))
  t$fips <- substr(t$site,1,5)
  t <- subset(t[order(t$fips,t$dv,decreasing=TRUE),],!duplicated(fips))
  table4 <- data.frame(state_name=t$state_name,county_name=t$county_name,
    state_fips=substr(t$fips,1,2),county_fips=substr(t$fips,3,5),
    epa_region=t$epa_region,site=t$site,dv=t$dv,cbsa_name=t$cbsa_name)
  table4 <- table4[order(table4$state_name,table4$county_name),]
  writeData(wb,sheet=4,x=table4,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=4,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=4,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=4,years); set.footnote.dates(wb,sheet=4);
  remove.extra.rows(wb,sheet=4,row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5. Site Status
  t <- subset(dvs,dv_year == year)
  t$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv," "),t$dv,t$valid)
  t$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "N",dv," "),t$dv,t$valid)
  table5 <- t[,c("state_name","county_name","cbsa_name","csa_name","naa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv","valid_months",
    "dv_month","valid.yr1","valid.yr2","valid.yr3","max.yr1","max.yr2","max.yr3","month.yr1",
    "month.yr2","month.yr3")]
  writeData(wb,sheet=5,x=table5,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=5,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=5,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=5,years); set.footnote.dates(wb,sheet=5);
  remove.extra.rows(wb,sheet=5,row.hts=c(15,47,15,47,15,79,15))
  
  ## Table 6. Site Trends
  t <- dcast(subset(dvs,valid == "Y"),site ~ dv_year,value.var="dv")
  colnames(t)[2:11] <- paste("dv",years[1:10],years[3:12],sep="_")
  vals <- merge(subset(dvs,!duplicated(site)),t,by="site")
  table6 <- vals[,c("state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region","site","site_name","address","latitude","longitude",
     paste("dv",years[1:10],years[3:12],sep="_"))]
  writeData(wb,sheet=6,x=table6,startCol=1,startRow=5,colNames=FALSE,na.string="")
  writeData(wb,sheet=6,x=paste("AQS Data Retrieval:",today),startCol=1,startRow=2,colNames=FALSE)
  writeData(wb,sheet=6,x=paste("Last Updated:",today),startCol=3,startRow=2,colNames=FALSE)
  set.year.values(wb,sheet=6,years); set.footnote.dates(wb,sheet=6);
  remove.extra.rows(wb,sheet=6,row.hts=c(15,31,15,47,15,79,15))
  
  ## Write DV tables to Excel File
  fix.scripts(wb)
  file.xlsx <- paste("Pb_DesignValues",(year-2),year,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  saveWorkbook(wb,file=paste("DVs2xlsx/",year,"/",file.xlsx,".xlsx",sep=""),overwrite=TRUE)
}