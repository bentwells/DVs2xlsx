## Format NO2 DV spreadsheets based on template file
no2.dvtables <- function(years,step1.date,type="DRAFT") {
  
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
  require(plyr); require(xlsx);
  ny <- length(years); dvyr1 <- years[(ny-2)]; dvyr2 <- years[(ny-1)]; dvyr3 <- years[ny];
  dv.begin.date <- as.Date(paste(dvyr1,"01-01",sep="-"))
  work.dir <- paste("NO2_DVs/DV",dvyr1,"_",dvyr3,"/",step1.date,"/",sep="")
  out.dir <- paste("DVs2xlsx/",dvyr3,"/",sep="")
  
  ## Get template file and sheets
  templates <- list.files("DVs2xlsx/templates")
  template.file <- templates[intersect(grep("NO2",templates),grep(type,templates))]
  dv.wb <- loadWorkbook(paste("DVs2xlsx/templates",template.file,sep="/"))
  dv.sheets <- getSheets(dv.wb)
  
  ## Load site-level DVs (created by dvs_step1.r) 
  load(paste(work.dir,"dv_",years[3],"_",years[ny],"_ARC53.Rdata",sep=""))
  load(paste(work.dir,"dv_",years[1],"_",years[ny],"_HRC100.Rdata",sep=""))
  
  ## Retrieve nonattainment area info from AQS
  naa.info <- get.naa.info(par=42602,psid=8)
  naa.states <- get.naa.states(naa.info)
  
  ## Table 0: Monitor metadata
  monitors <- get.monitors(par=42602,yr1=years[1],yr2=dvyr3,all=TRUE)
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
    v_area <- ifelse(as.Date(mbd) >= as.Date(paste(dvyr3,"01-01",sep="-")) &
      base_info$state_name %in% naa.states,"1"," ")
    v_nonreg <- ifelse(nbd != " " & nrc == " " & as.Date(lsd) >= dv.begin.date,"1"," ")
    v_nonref <- ifelse(non_ref != " " & nrc != "Y" & check.dates == TRUE &
      as.Date(lsd) >= dv.begin.date,"1"," ")
    v_closure <- ifelse(med == " " & nrc != "Y" & as.Date(lsd) >= dv.begin.date &
      as.Date(lsd) <= as.Date(paste(dvyr3,"10-01",sep="-")),"1"," ")
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
  
  ## Table 1a: Nonattainment area status for the 1971 Annual NO2 NAAQS
  temp <- subset(dv.ann,naa_name != " " & eval(parse(text=paste("valid",dvyr3,sep="."))) == "Y",
    c("site","naa_name",paste("dv",dvyr3,sep=".")))
  table1a <- naa.info[,c("naa_name","epa_regions","status")]
  table1a$dv <- NA; table1a$meets_naaqs <- "Incomplete";
  if (nrow(temp) > 0) {
    table1a$dv <- tapply(temp[,paste("dv",dvyr3,sep=".")],list(temp$naa_name),max.na)
    table1a$meets_naaqs <- sapply(table1a$dv,function(x) ifelse(x > 53,"No","Yes"))
  }
  S <- ifelse(type == "DRAFT",2,1)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table1a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=8,col=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=12,col=1)
  
  ## Table 2a: Additional monitors violating the 1971 Annual NO2 NAAQS
  table2a <- subset(dv.ann,naa_name != " " & eval(parse(text=paste("dv",dvyr3,sep="."))) > 53 &
    eval(parse(text=paste("valid",dvyr3,sep="."))) == "Y",c("state_name","county_name",
    "epa_region","site",paste("dv",dvyr3,sep="."),"cbsa_name"))
  if (nrow(table2a) == 0) {
    table2a <- data.frame(x=paste("There were no sites violating the annual NO2 NAAQS in ",
      dvyr3,".",sep=""))
  }
  S <- ifelse(type == "DRAFT",3,2)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table2a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
  ## Table 2b: Monitors violating the 2010 1-hour NO2 NAAQS
  table2b <- subset(dv.1hr,eval(parse(text=paste("dv",dvyr1,dvyr3,sep="."))) > 100 & 
    eval(parse(text=paste("valid",dvyr1,dvyr3,sep="."))) == "Y",c("state_name","county_name",
    "epa_region","site",paste("dv",dvyr1,dvyr3,sep="."),"cbsa_name"))
  if (nrow(table2b) == 0) {
    table2b <- data.frame(x=paste("There were no sites violating the 1-hour NO2 NAAQS in ",
      dvyr1,"-",dvyr3,".",sep=""))
  }
  S <- ifelse(type == "DRAFT",4,3)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table2b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
  ## Table 3a: Nonattainment area trends for the 1971 Annual NO2 NAAQS
  temp <- subset(dv.ann,naa_name != " ")
  temp[,grep("dv",colnames(temp))] <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),
    temp[,grep("dv",colnames(temp))],temp[,grep("valid",colnames(temp))])
  vals <- apply(temp[,grep("dv",colnames(temp))],2,function(x) 
    tapply(x,list(temp$naa_name),max.na))
  table3a <- data.frame(naa.info[,c("naa_name","epa_regions")],t(vals))
  S <- ifelse(type == "DRAFT",5,4)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=2)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table3a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=10,col=1)
  
  ## Table 4a: County-level design values for the 1971 Annual NO2 NAAQS
  temp <- subset(dv.ann,eval(parse(text=paste("valid",dvyr3,sep="."))) == "Y",c("site",
    "state_name","county_name","epa_region",paste("dv",dvyr3,sep="."),"cbsa_name"))
  temp <- temp[order(temp$state_name,temp$county_name,temp$dv,decreasing=TRUE),]
  temp <- temp[!duplicated(paste(temp$state_name,temp$county_name)),]
  temp <- temp[order(temp$state_name,temp$county_name),]
  table4a <- data.frame(state_name=temp$state_name,county_name=temp$county_name,
    state_fips=substr(temp$site,1,2),county_fips=substr(temp$site,3,5),
    epa_region=temp$epa_region,site=temp$site,dv=temp$dv,cbsa_name=temp$cbsa_name)
  S <- ifelse(type == "DRAFT",6,5)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table4a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=259,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 4b: County-level design values for the 2010 1-hour NO2 NAAQS
  temp <- subset(dv.1hr,eval(parse(text=paste("valid",dvyr1,dvyr3,sep="."))) == "Y",c("site",
    "state_name","county_name","epa_region",paste("dv",dvyr1,dvyr3,sep="."),"cbsa_name"))
  temp <- temp[order(temp$state_name,temp$county_name,temp$dv,decreasing=TRUE),]
  temp <- temp[!duplicated(paste(temp$state_name,temp$county_name)),]
  temp <- temp[order(temp$state_name,temp$county_name),]
  table4b <- data.frame(state_name=temp$state_name,county_name=temp$county_name,
    state_fips=substr(temp$site,1,2),county_fips=substr(temp$site,3,5),
    epa_region=temp$epa_region,site=temp$site,dv=temp$dv,cbsa_name=temp$cbsa_name)
  S <- ifelse(type == "DRAFT",7,6)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table4b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=229,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,63,15))
  
  ## Table 5a: Site-level design values for the 1971 Annual NO2 NAAQS
  temp <- subset(dv.ann,!is.na(eval(parse(text=paste("valid",dvyr3,sep=".")))),c("site","site_name",
    "address","latitude","longitude","state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region",paste("dv",dvyr3,sep="."),paste("pct",dvyr3,sep="."),paste("valid",dvyr3,sep=".")))
  temp$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),
    dv=temp[,paste("dv",dvyr3,sep=".")],valid=temp[,paste("valid",dvyr3,sep=".")])
  temp$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",NA,dv),
    dv=temp[,paste("dv",dvyr3,sep=".")],valid=temp[,paste("valid",dvyr3,sep=".")])
  table5a <- temp[,c("state_name","county_name","cbsa_name","csa_name","naa_name",
    "epa_region","site","site_name","address","latitude","longitude","valid_dv",
    "invalid_dv",paste("pct",dvyr3,sep="."))]
  S <- ifelse(type == "DRAFT",8,7)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table5a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=459,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 5b: Site-level design values for the 2010 1-hour NO2 NAAQS
  temp <- subset(dv.1hr,!is.na(eval(parse(text=paste("valid",dvyr1,dvyr3,sep=".")))),c("site",
    "site_name","address","latitude","longitude","state_name","county_name","cbsa_name","csa_name",
    "epa_region",paste("dv",dvyr1,dvyr3,sep="."),paste("valid",dvyr1,dvyr3,sep="."),
    paste("p98",dvyr1:dvyr3,sep="."),paste("pct.q",rep(1:4,3),".",rep(dvyr1:dvyr3,each=4),sep="")))
  temp$valid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),
    dv=temp[,paste("dv",dvyr1,dvyr3,sep=".")],valid=temp[,paste("valid",dvyr1,dvyr3,sep=".")])
  temp$invalid_dv <- mapply(function(dv,valid) ifelse(valid == "Y",NA,dv),
    dv=temp[,paste("dv",dvyr1,dvyr3,sep=".")],valid=temp[,paste("valid",dvyr1,dvyr3,sep=".")])
  temp$qtrs.yr1 <- apply(temp[,paste("pct.q",c(1:4),".",dvyr1,sep="")],1,function(x) sum(x >= 75))
  temp$qtrs.yr2 <- apply(temp[,paste("pct.q",c(1:4),".",dvyr2,sep="")],1,function(x) sum(x >= 75))
  temp$qtrs.yr3 <- apply(temp[,paste("pct.q",c(1:4),".",dvyr3,sep="")],1,function(x) sum(x >= 75))
  table5b <- temp[,c("state_name","county_name","cbsa_name","csa_name","epa_region",
    "site","site_name","address","latitude","longitude","valid_dv","invalid_dv",
    "qtrs.yr1","qtrs.yr2","qtrs.yr3",paste("p98",dvyr1:dvyr3,sep="."))]
  S <- ifelse(type == "DRAFT",9,8)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table5b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6a: Trends in site-level design values for the 1971 Annual NO2 NAAQS
  dv.ann[,grep("dv",colnames(dv.ann))] <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),
    dv.ann[,grep("dv",colnames(dv.ann))],dv.ann[,grep("valid",colnames(dv.ann))])
  counts <- apply(dv.ann[,grep("dv",colnames(dv.ann))],1,function(x) sum(!is.na(x)))
  table6a <- dv.ann[which(counts > 0),c("state_name","county_name","cbsa_name","csa_name",
    "naa_name","epa_region","site","site_name","address","latitude","longitude",
    colnames(dv.ann)[grep("dv",colnames(dv.ann))])]
  S <- ifelse(type == "DRAFT",10,9)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table6a,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=589,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Table 6b: Trends in site-level design values for the 2010 1-hour NO2 NAAQS
  dv.1hr[,grep("dv",colnames(dv.1hr))] <- mapply(function(dv,valid) ifelse(valid == "Y",dv,NA),
    dv.1hr[,grep("dv",colnames(dv.1hr))],dv.1hr[,grep("valid",colnames(dv.1hr))])
  counts <- apply(dv.1hr[,grep("dv",colnames(dv.1hr))],1,function(x) sum(!is.na(x)))
  table6b <- dv.1hr[which(counts > 0),c("state_name","county_name","cbsa_name","csa_name",
    "epa_region","site","site_name","address","latitude","longitude",
    colnames(dv.1hr)[grep("dv",colnames(dv.1hr))])]
  S <- ifelse(type == "DRAFT",11,10)
  set.aqs.date(dv.sheets[[S]],step1.date,row=2,col=1)
  set.last.update(dv.sheets[[S]],row=2,col=3)
  set.year.values(dv.sheets[[S]],years,row=4)
  df2xls(df=table6b,sheet=dv.sheets[[S]],sr=5,sc=1)
  set.footnote.date(dv.sheets[[S]],step1.date,row=509,col=1)
  remove.extra.rows(dv.wb,dv.sheets[[S]],row.hts=c(15,31,15,47,15,79,15))
  
  ## Save DV tables in .Rdata format and write to Excel file
  file.rdata <- paste("dvtables",dvyr1,dvyr3,format(Sys.Date(),"%Y%m%d"),sep="_")
  file.xlsx <- paste("NO2_DesignValues",dvyr1,dvyr3,type,format(Sys.Date(),"%m_%d_%y"),sep="_")
  save(list=ls(pattern="table"),file=paste(work.dir,file.rdata,".Rdata",sep=""))
  saveWorkbook(dv.wb,file=paste(out.dir,file.xlsx,".xlsx",sep=""))
}
