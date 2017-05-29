#' Этот пакет позволяет импортировать данные с сайта Росстата (http://www.gks.ru).
#' @name RGksStatData
#' @docType package
#' @import XML
#' @import xml2
#' @import tools
NULL

#' Запускает функции, которые позволяют получить данные в виде фрейма, в зависимости от типа источника
#' @param 
#' ref - ссылка на .doc файл, на сайте Росстата, результат функции getGKSDataRef()
#' @examples
#' loadGKSData(getGKSDataRef())
#' @export
loadGKSData <- function(ref){
    ext <- file_ext(ref)
    if(ext == "doc"){
        ref <- gsub("/%3Cextid%3E/%3Cstoragepath%3E::\\|","",ref)
        docx_fail <- Transform_doc_to_docx(ref)
        getTableFromDocx(docx_fail)
        unlink(docx_fail)
    }else if(ext == "htm"){
        getTableFromHtm(ref)
    }else
        print("Неподдерживаемый тип источника")
    
}

#' Исправляет кодировку для загруженного фрейма данных
#' @param 
#' x - data.frame с данными
#' @param 
#' sep - разделитель полей в csv файле
#' @param 
#' quote - экранирование данных кавычками
#' @param 
#' encoding - нужная кодировка
#' @examples
#' toLocalEncoding(dataGKS)
#' @export
toLocalEncoding <- function(x, sep=",", quote=TRUE, encoding="utf-8"){
    rawcsv <- tempfile()
    write.csv(x, file = rawcsv)
    result <- read.csv(rawcsv, encoding = "UTF-8")
    unlink(rawcsv)
    result
}

#' Отображает список доступной статистики Росстата, позволяет производить навигацию по списку и получить ссылку на выбранный документ
#' @return 
#' Расположение файла на сайте Росстата
#' @examples
#' getGKSDataRef()
#' @export
getGKSDataRef <- function(){
    path <- '/bgd/regl/'
    params <- '/?List&Id='
    id <- -1
    year <- readline(prompt = paste("Введите год от", years$years_v[1],
                                    "до", tail(years$years_v, n = 1), " "))
    db_name <- as.character(years$db_names[years_v == year])
    while(TRUE){
        url <- paste(host_name, path, db_name, params, id, sep = '')
        download.file(url, destfile = 'test_xml.xml')
        test <- readLines('test_xml.xml')
        xml <- xmlTreeParse(url, useInternalNodes = T)
        names <- xpathSApply(xml, "//name", xmlValue)
        Encoding(names) <- "UTF-8" 
        refs <- xpathSApply(xml, "//ref", xmlValue)
        for(i in 1:length(names))
            print(paste(i, names[i]))
        num <- readline(prompt = "Введите номер ")
        ref <- refs[as.numeric(num)]
        if(substr(ref, 1, 1) != "?")
            return(ref)
        id <- substr(ref, 2, nchar(ref))
    }
}

#' Позволяет получить данные в виде фрейма из htm/html страницы, где есть 1 таблица с данными
#' @param 
#' word_doc - ссылка на htm/html страницу, на сайте Росстата
#' @return 
#' В переменную dataGKS записывает данные в виде фрейма
#' @example 
#' getTableFromHtm(getGKSDataRef())
#' 2006
#' 3
#' 2
#' @export
getTableFromHtm <- function(ref) {
    url <- paste(host_name, ref, sep = "")
    doc <- htmlParse(url, encoding = "Windows-1251")
    if(length(xpathSApply(doc,"//table", xmlValue)) == 0){
        print("Нет таблицы в источнике")
        return()
    }
    
    assign("dataGKS", readHTMLTable(doc, trim = TRUE, which = 1,
                                    stringsAsFactors = FALSE,
                                    as.data.frame = TRUE), .GlobalEnv)
    names(dataGKS) <<- gsub("[\r\n]", "", names(dataGKS))
    dataGKS[, 1] <<- gsub("[\r\n]", "", dataGKS[, 1])
    dataGKS <<- toLocalEncoding(dataGKS)
}

#' Позволяет получить документ в формате .docx, из документа .doc, находящегося на сайте Росстата
#' @param 
#' word_doc - ссылка на .doc файл, на сайте Росстата
#' @return 
#' Расположение .docx файла
#' @examples
#' ref <- getGKSDataRef()
#' 2016
#' 3
#' 2
#' ref <- gsub("/%3Cextid%3E/%3Cstoragepath%3E::\\|","",ref)
#' Transform_doc_to_docx(ref)
#' @export
Transform_doc_to_docx <- function(word_doc){
  word_dir <- shell('reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE" /v PATH', intern = TRUE)
  winword_dir <- c(strsplit(word_dir[grep("PATH",word_dir)],"  ")[[1]])
  winword_dir <- winword_dir[length(winword_dir)]
  
  
  tmpd <- tempdir()
  tmpf2_1 <- tempfile(tmpdir=tmpd, fileext=".doc")
  download.file(paste(host_name, word_doc,sep=""),
                tmpf2_1, mode = "wb")
  kat <- getwd()
  tmpf2_2 <- tempfile(tmpdir=kat, fileext=".docx")
  tmpf2_2 <- gsub('\\\\','/',tmpf2_2)
  conv_to_docx <- paste('"',winword_dir, 'wordconv.exe" -oice -nme ',tmpf2_1, ' ',tmpf2_2, sep="")
  shell(conv_to_docx)
  
  unlink(tmpf2_1)
  return(tmpf2_2)
}

#' Позволяет получить данные в виде фрейма из документа в формате docx, где есть 1 таблица с данными
#' @param 
#' word_doc - расположение .docx файла
#' @return 
#' В переменную dataGKS записывает данные в виде фрейма
#' @example 
#' getTableFromDocx("C:/username/documents/file.docx")
#' @export
getTableFromDocx <- function(word_doc) {
    
  tmpd <- tempdir()
  tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")
  file.copy(word_doc, tmpf)
    
  unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))
    
  doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))
    
  unlink(tmpf)
  unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)
  
  ns <- xml_ns(doc)
  
  tbls <- xml_find_all(doc, ".//w:tbl", ns=ns)
  
  lapply_result <- lapply(tbls, function(tbl) {
      
      cells <- xml_find_all(tbl, "./w:tr/w:tc", ns=ns)
      rows <- xml_find_all(tbl, "./w:tr", ns=ns)
      dat <- data.frame(matrix(xml_text(cells), 
                               ncol=(length(cells)/length(rows)), 
                               byrow=TRUE), 
                        stringsAsFactors=FALSE)
      colnames(dat) <- dat[1,]
      dat <- dat[-1,]
      rownames(dat) <- NULL
      return(dat)
  })
  if (length(lapply_result)>1){
    dataGKS <- lapply_result[[1]]
    for (i in 2:length(lapply_result)){
      dataGKS <- rbind(dataGKS,lapply_result[[i]])
    }
    assign("dataGKS", dataGKS, .GlobalEnv)
  }else{
    assign("dataGKS", lapply_result, .GlobalEnv)
  }
    
}

#' Добавляет коды ОКАТО субъектов РФ к фрейму, содержащему панельные данные, необходим файл subjects.csv, который содержит коды ОКАТО
#' @param 
#' dann_frame_main - фрейм данных
#' @param id - индекс столбца, содержащего названия регионов, по умолчанию 1
#' @return исходный фрейм данных со столбцом OKATO
#' @examples
#' Add_OKATO(dataGKS)
#' @export
Add_OKATO <- function(dann_frame_main, id = 1){
    dann_frame <- dann_frame_main
    dann_frame[,id] <- toupper(dann_frame[,id])
    Subj_table <- read.table(file = system.file("extdata", "subjects.csv", package="RGksStatData"), sep=",",col.names=c("Full_name","OKATO","Short_name"))
    OKATO <- c()
    for (i in 1:nrow(dann_frame[id])){
        
        ma <- FALSE
        fed_okr <- grepl("ФЕДЕРАЛЬНЫЙ",as.character(dann_frame[i,id]))||
          grepl("ДОЛГАНО-НЕНЕЦКИЙ",as.character(dann_frame[i,id]))||
          grepl("КОМИ-ПЕРМЯЦКИЙ",as.character(dann_frame[i,id]))
        if (!fed_okr)
            for (j in 1:length(Subj_table$Short_name)){
                if (as.character(Subj_table$Short_name[j])=='НЕНЕЦКИЙ'){
                    yane <- grepl('ЯМАЛО-НЕНЕЦКИЙ',as.character(dann_frame[i,id]))
                    if (yane){
                        ma <- FALSE
                    }else ma <- grepl(as.character(Subj_table$Short_name[j]),as.character(dann_frame[i,id]))
                }else ma <- grepl(as.character(Subj_table$Short_name[j]),as.character(dann_frame[i,id]))
                
                if (ma){
                    OKATO <- c(OKATO, as.character(Subj_table$OKATO[j]))
                    break;
                }
            }
        if (!ma){
            OKATO <- c(OKATO, NA)
        }
    }
    
    new_dann_frame <- data.frame(dann_frame_main,OKATO)
    return(new_dann_frame)
}

#' Возвращает данные в виде временного ряда по региону
#' @param 
#' dann - фрейм данных, c кодами ОКАТО в последнем столбце
#' @param ind_name - название показателя
#' @param id - индекс столбца, содержащего названия регионов, по умолчанию 1
#' @return 
#' Список: 
#' [[1]] - Фрейм, содержащий название региона и код ОКАТО
#' [[2]] - Название показателя
#' [[3]] - Фрейм, содержащий года и значения показателя по годам
#' @examples
#' info_region(Add_OKATO(dataGKS), ind_name = "Население")
#' @export
info_region <- function(dann,ind_name = "",id = 1){
  show("Регионы:")
  show(dann[,id])
  region <- readline(prompt = paste("Введите название региона ")) 
  index_region <- 0
  
  for (i in 1:nrow(dann)){
    gre <- grepl(region,as.character(dann[i,1]))
    if (gre){
      index_region <- i
      break;
    }
  }
  
  if (index_region==0){
    show("Регион не найден")
    return()
  }else{
    dann_region <- dann[index_region,]
    year <- colnames(dann_region)[-(1:id)]
    year <- year[-length(year)]
    Value <- c()
    for (i in (id+1):(ncol(dann_region)-1)){
      Value <- c(Value,as.vector(dann_region[1,i]))
    }
    new_dann <- data.frame(year,Value)
    
    l.1 <- data.frame(Region = region, OKATO = as.vector(dann_region[1,ncol(dann_region)]))
    l.2 <- ind_name
    l.3 <- data.frame(year,Value)
    
    return(list(l.1,l.2,l.3))
  }
}

#' Возвращает данные из сборника Росстата в виде временного ряда по региону или в виде панельных данных с кодами ОКАТО 
#' @param 
#' year - год сборника
#' @return 
#' фрейм данных с кодами ОКАТО или список, результат функции info_region
#' @examples
#' Return_data()
#' @export
Return_data <- function(year = 2016){
  db_name <- as.character(years$db_names[years_v == year])
  path <- '/bgd/regl/'
  params <- '/?List&Id='
  id <- -1
  while(TRUE){
    url <- paste(host_name, path, db_name, params, id, sep = '')
    download.file(url, destfile = 'test_xml.xml')
    test <- readLines('test_xml.xml')
    xml <- xmlTreeParse(url, useInternalNodes = T)
    names <- xpathSApply(xml, "//name", xmlValue)
    Encoding(names) <- "UTF-8" 
    refs <- xpathSApply(xml, "//ref", xmlValue)
    for(i in 1:length(names))
      print(paste(i, names[i]))
    num <- readline(prompt = "Введите номер ")
    num <- as.numeric(num)
    ind_name <- names[num]
    ref <- refs[num]
    if(substr(ref, 1, 1) != "?")
      break;
    id <- substr(ref, 2, nchar(ref))
  }
  loadGKSData(ref)
  mode <- readline(prompt = paste("Выберите вид выходных данных:\n 1 - панельные данные\n 2 - временной ряд по региону/федеральному округу ")) 
  mode <- as.numeric(mode)
  if (mode==1){
    new_dann <- Add_OKATO(dataGKS,1)
    return(new_dann)
  }else
  if (mode==2){
    new_dann <- Add_OKATO(dataGKS,1)
    return(info_region(new_dann, ind_name))
    
  }else{
    show("Такой вид выходных данные отсутствует")
    return()
  }
}

