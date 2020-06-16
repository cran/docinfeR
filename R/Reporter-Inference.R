library(officer)
library(nortest)
library(flextable)
library(broom)
library(tictoc)

tictoc::tic("Global")
# Inference Report
my.OS <- Sys.info()['sysname']



main <- function(path, var.type="all"){

#Read .csv File

if(path == ""){
	path <- file.choose(new = FALSE)
}
file_in <- file(path,"r")
#file_out <- file(out.path <<- paste(path,"-REPORT-INFERENCE-",var.type,".docx",sep = ""),open="wt", blocking = FALSE, encoding ="UTF-8")
file_out <- file(paste(tempdir(),"docinfeR-REPORT-INFERENCE-",var.type,"-",format(Sys.time(), "%a-%b-%d-%X-%Y"),".docx",sep = ""),open="wt", blocking = FALSE, encoding ="UTF-8")

#defining the document name
if(grepl("linux",tolower(my.OS))){
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-INFERENCE-",var.type,".docx",sep = "")
  aux.file.report <- file_out
  dir.create("/tmp/data/", showWarnings = FALSE)
  aux.file.rdata <- paste("/tmp/data/", my.time,"-TEMP.DATA",sep="")
}else{
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-INFERENCE-",var.type,".docx",sep = "")
  dir.create("/tmp/data/", showWarnings = FALSE)
  aux.file.rdata <- paste("/tmp/data/", my.time,"-TEMP.DATA",sep="")
}
my_doc <- officer::read_docx()


factor.factor <- function(vars.val.pass,var1type,var2type){
  if(length(unique(vars.val.pass[,var1type]))==1 || length(unique(vars.val.pass[,var2type]))==1){
    warning(Sys.time()," - var1:",paste(var1type)," - var2:",paste(var2type)," - [ERROR][INF] Error in stats::chisq.test: 'x' must at least have 2 elements\n\n")
    return({0})
  }else{
    
    crosstbl <- table(vars.val.pass[,var1type],vars.val.pass[,var2type],dnn=c(var1type,var2type))
    crosstbl_matrix <-as.matrix(crosstbl)
    count_1<-which(crosstbl_matrix<5)
    count_2<-which(crosstbl_matrix>=5)
    ratio1 <- length(count_1)/(nrow(crosstbl_matrix)*ncol(crosstbl_matrix))
    ratio2 <-length(count_2)/(nrow(crosstbl_matrix)*ncol(crosstbl_matrix))
    
    #ATTENTION: Crosstable can get big!!
    if(nrow(crosstbl_matrix)*ncol(crosstbl_matrix)<100){
      val.test.inf <- flextable::flextable(as.data.frame.table(crosstbl_matrix))
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }else{
      officer::body_add_par(my_doc, paste("NOTE: TEST Crosstable of high dimension. Therefore, not available in this report!"),style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
    
    
    #if there are zeros and 80% or less of the values are less than 5
    #do Fisher exact test
    if(any(crosstbl == 0 )){
      if(ratio1>0.20 && ratio1 < 0.80){
        
        
        val.test.inf <- flextable::flextable(broom::tidy(my.fisher <- stats::fisher.test(x = crosstbl, alternative = "two.sided",simulate.p.value = TRUE)))
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste("The results of the analysis shows that Fisher test presents a p-value of ",round(my.fisher$p.value,3)," ",ifelse(my.fisher$p.value<0.05,paste("(p<0.05)."),paste("(p>0.05)."))," Thus, the null hipothesis is ",ifelse(my.fisher$p.value<0.05,paste("rejected"),paste("non-rejected"))," with a confidence level of 95%. We can conclude that both variables are ",ifelse(my.fisher$p.value<0.05,paste("dependent."),paste("independent.")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        
        officer::body_add_par(my_doc, paste("NOTE: TEST not robust due to high percentage of values inferior to 5 in crosstable: ",ratio1*100,"%\n\n",sep=""), style = "Normal",pos="after") 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }else if(ratio1 < 0.20){
        val.test.inf <- flextable::flextable(broom::tidy(my.fisher <- stats::fisher.test(x = crosstbl, alternative = "two.sided",simulate.p.value = TRUE)))
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste("The results of the analysis shows that Fisher test presents a p-value of ",round(my.fisher$p.value,3)," ",ifelse(my.fisher$p.value<0.05,paste("(p<0.05)."),paste("(p>0.05)."))," Thus, the null hipothesis is ",ifelse(my.fisher$p.value<0.05,paste("rejected"),paste("non-rejected"))," with a confidence level of 95%. We can conclude that both variables are ",ifelse(my.fisher$p.value<0.05,paste("dependent."),paste("independent.")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }else{
        return({0})
      } 
    }else{
      if(ratio1>0.20 && ratio1 < 0.80){
        val.test.inf <- flextable::flextable(broom::tidy(my.chisq <- stats::chisq.test(x = crosstbl)))####what happens if there are values in the crosstbl equal to 5
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste("The results of the analysis shows that chi-squared test presents a p-value of ",round(my.chisq$p.value,3)," ",ifelse(my.chisq$p.value<0.05,paste("(p<0.05)."),paste("(p>0.05)."))," Thus, the null hipothesis is ",ifelse(my.chisq$p.value<0.05,paste("rejected"),paste("non-rejected"))," with a confidence level of 95%. We can conclude that both variables are ",ifelse(my.chisq$p.value<0.05,paste("dependent."),paste("independent.")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        officer::body_add_par(my_doc, paste("NOTE: TEST not robust due to high percentage of values inferior to 5 in crosstable: ",ratio1*100,"%\n\n",sep=""), style = "Normal",pos="after") 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }else if(ratio1 < 0.20){
        val.test.inf <- flextable::flextable(broom::tidy(my.chisq <- stats::chisq.test(x = crosstbl)))####what happens if there are values in the crosstbl equal to 5
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste("The results of the analysis shows that chi-squared test presents a p-value of ",round(my.chisq$p.value,3)," ",ifelse(my.chisq$p.value<0.05,paste("(p<0.05)."),paste("(p>0.05)."))," Thus, the null hipothesis is ",ifelse(my.chisq$p.value<0.05,paste("rejected"),paste("non-rejected"))," with a confidence level of 95%. We can conclude that both variables are ",ifelse(my.chisq$p.value<0.05,paste("dependent."),paste("independent.")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }else{
        return({0})
      } 
    }
  }
}

factor.numeric.integer <- function(vars.val.pass,var1type,var2type){
  #check normality for quantitative variable
  #with Kolmogorov-Smirnov (lilliefors)
  if(length(unique(stats::na.omit(vars.val.pass[,var2type])))<2 || length(stats::na.omit(vars.val.pass[,var2type]))<5){
    return({0})
  }else{
    p.value <-nortest::lillie.test(stats::na.omit(vars.val.pass[,var2type]))$p.value###The K-S test is for a continuous distribution and so MYDATA
  }
  
  aux.data <- vars.val.pass[stats::complete.cases(vars.val.pass[,c(var1type,var2type)]),c(var1type,var2type)]
  
  #should not contain any ties (repeated values). 
  #https://stats.stackexchange.com/questions/232011/ties-should-not-be-present-in-one-sample-kolmgorov-smirnov-test-in-r
  if(p.value < 0.05){#not Normal
    
    
    if(length(unique(aux.data[,var2type]))==1){
      warning(Sys.time()," - var2:",paste(var2type)," - [ERROR][INF] Mann-Whitney-Wilcoxon or Kruskal cannot be applied to numeric vectors with only 1 unique value\n\n")
      return({0})
    }
    if(length(unique(aux.data[,var1type]))==2){###What do you mean with unique(var1)==2?
      
      #Mann-Whitney
      test <-  broom::tidy(my.wilcox <- stats::wilcox.test(aux.data[,var2type]~aux.data[,var1type]))  #stats::wilcox.test(y~A) where y is numeric and A is A binary factor (independent)
      test$data.name<-paste(var2type,"by",var1type,sep = " ")
      val.test.inf <- flextable::flextable(test)
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste('Analyses show that p-value is ',round(my.wilcox$p.value,3),ifelse(round(my.wilcox$p.value,3)<0.05,paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups of "',var2type,'" are very different regarding "',var1type,'" variable.',sep=""),paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, the distributions are similar. Thus, it is possible to assume that groups of "',var2type,'" are similar regarding "',var1type,'" variable.',sep="")),sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }else if(length(unique(aux.data[,var1type]))>2){
      
      #Kruskal-Wallis
      if(length(unique(aux.data[,var1type]))==1 || length(unique(aux.data[,var2type]))<=2){
        warning(Sys.time()," - var1:",paste(var1type)," - [ERROR][INF] Error in stats::kruskal.test.default: all observations are in the same group\n\n")
        return({0})
      }else{
        test <- broom::tidy(my.kruskal <- stats::kruskal.test(x=aux.data[,var2type],g=aux.data[,var1type]))#https://www.statmethods.net/stats/nonparametric.html #stats::kruskal.test(y~A) # where y1 is numeric and A is a factor
        test$data.name<-paste(var2type,"by",var1type,sep = " ")
        val.test.inf <- flextable::flextable(test)
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.kruskal$p.value,3),ifelse(round(my.kruskal$p.value,3)<0.05,paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups are very different regarding "',var1type,'" variable.',sep=""),paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups have similar distribution regarding "',var1type,'" variable.',sep="")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }
    }else{
      return({0})
    }
  }else{#normal
    vars.val.pass <- stats::na.omit(vars.val.pass)
    
    if(length(unique(vars.val.pass[,var1type]))==2){
      
      #ttest
      tryCatch(stop(val.test.inf <- flextable::flextable(broom::tidy(my.t.test <- stats::t.test(vars.val.pass[,var2type]~vars.val.pass[,var1type], var.equal=TRUE)))), error = function(e) {warning(Sys.time()," - var1:",paste(var1type)," - var2:",paste(var2type)," - [ERROR][INF] ",e,"\n\n")}, finally = {
        # stats::t.test(y~x) where y is numeric and x is a binary factor
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.t.test$p.value,3),' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups of "',var1type,'" have the same distribution regarding variable "',var2type,'".',sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      })
      
    }else{##What do you mean with unique(var1)>2?
      #Anova
      if(length(unique(vars.val.pass[,var1type]))<2){
        warning(Sys.time()," - var1:",paste(var1type)," - [ERROR][INF] Error in contrasts<-: contrasts can be applied only to factors with 2 or more levels\n\n")
        return({0})
      }else{
        test<- broom::tidy(my.aov <- stats::aov(vars.val.pass[,var2type] ~ vars.val.pass[,var1type]))     #"Anova"
        my.aov$p.value <- stats::summary.aov(my.aov)[[1]][["Pr(>F)"]][1]
        
        if(is.null(my.aov$p.value)){
          return({0})
        }
        
        test$call <- paste("stats::aov(formula = ",var2type," ~ ",var1type,")",sep = "")
        val.test.inf <- flextable::flextable(test)
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.aov$p.value,3),ifelse(round(my.aov$p.value,3)>=0.05,paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups have the same distribution regarding "',var2type,'" variable.',sep=""),paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups do not have the same distribution regarding "',var2type,'" variable.',sep="")),sep = "") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }
    }
  }
}

numeric.integer.factor <- function(vars.val.pass,var1type,var2type){
  
  #check normality for quantitative variable
  #with Kolmogorov-Smirnov (lilliefors)
  if(length(unique(stats::na.omit(vars.val.pass[,var1type])))<2 || length(stats::na.omit(vars.val.pass[,var1type]))<5){
    return({0})
  }else{
    p.value <-nortest::lillie.test(stats::na.omit(vars.val.pass[,var1type]))$p.value###The K-S test is for a continuous distribution and so MYDATA
  }
  
  aux.data <- vars.val.pass[stats::complete.cases(vars.val.pass[,c(var1type,var2type)]),c(var1type,var2type)]
  
  
  #should not contain any ties (repeated values). 
  #https://stats.stackexchange.com/questions/232011/ties-should-not-be-present-in-one-sample-kolmgorov-smirnov-test-in-r
  if(p.value < 0.05){#not Normal
    
    
    
    if(length(unique(aux.data[,var2type]))==2){
      
      #Mann-Whitney
      test <-  broom::tidy(my.wilcox <- stats::wilcox.test(aux.data[,var1type]~aux.data[,var2type]))  #stats::wilcox.test(y~A) where y is numeric and A is A binary factor (independent)
      test$data.name<-paste(var1type,"by",var2type,sep = " ")
      val.test.inf <- flextable::flextable(test)
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste('Analyses show that p-value is ',round(my.wilcox$p.value,3),ifelse(round(my.wilcox$p.value,3)<0.05,paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups of "',var2type,'" are very different regarding "',var1type,'" variable.',sep=""),paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, the distributions are similar. Thus, it is possible to assume that groups of "',var2type,'" are similar regarding "',var1type,'" variable.',sep="")),sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }else if(length(unique(aux.data[,var2type]))>2){
      #Kruskal-Wallis
      if(length(unique(aux.data[,var2type]))==1 || length(unique(aux.data[,var1type]))==1){
        warning(Sys.time()," - var2:",paste(var2type)," - [ERROR][INF] Error in stats::kruskal.test.default: all observations are in the same group\n\n")
        return({0})
      }else{
        test <- broom::tidy(my.kruskal <- stats::kruskal.test(aux.data[,var1type]~aux.data[,var2type]))#https://www.statmethods.net/stats/nonparametric.html #stats::kruskal.test(y~A) # where y1 is numeric and A is a factor
        test$data.name<-paste(var1type,"by",var2type,sep = " ")
        val.test.inf <- flextable::flextable(test)
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.kruskal$p.value,3),ifelse(round(my.kruskal$p.value,3)<0.05,paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups are very different regarding "',var1type,'" variable.',sep=""),paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups have similar distribution regarding "',var1type,'" variable.',sep="")),sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }
    }else{
      return({0})
    }
  }else{#normal
    #take out NAs
    vars.val.pass <- stats::na.omit(vars.val.pass)
    if(length(unique(vars.val.pass[,var2type]))==2){###What do you mean with unique(var1)==2?
      #ttest
      tryCatch(stop(val.test.inf <- flextable::flextable(broom::tidy(my.t.test <- stats::t.test(vars.val.pass[,var1type]~vars.val.pass[,var2type], var.equal=TRUE)))), error = function(e) {warning(Sys.time()," - var1:",paste(var1type)," - var2:",paste(var2type)," - [ERROR][INF] ",e,"\n\n")}, finally = {
        # stats::t.test(y~x) where y is numeric and x is a binary factor
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.t.test$p.value,3),' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups of "',var1type,'" have the same distribution regarding variable "',var2type,'".',sep="") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      })
    }else if(length(unique(vars.val.pass[,var2type]))>2){##What do you mean with unique(var1)>2?
      #Anova
      if(length(unique(vars.val.pass[,var2type]))==1){
        warning(Sys.time()," - ",paste(var2type)," - [ERROR][INF] Error in contrasts<-: contrasts can be applied only to factors with 2 or more levels\n\n")
        return({0})
      }else{
        
        test<- broom::tidy(my.aov <- stats::aov(vars.val.pass[,var1type] ~ vars.val.pass[,var2type]))     #"Anova"
        my.aov$p.value <- stats::summary.aov(my.aov)[[1]][["Pr(>F)"]][1]
        
        if(is.null(my.aov$p.value)){
          return({0})
        }
        
        test$call <- paste("stats::aov(formula = ",var1type," ~ ",var2type,")",sep = "")
        val.test.inf <- flextable::flextable(test)
        flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
        aux.text.inf <- paste('Analyses show that p-value is ',round(my.aov$p.value,3),ifelse(round(my.aov$p.value,3)>=0.05,paste(' (p>0.05). This means that null hypothesis is non-rejected (p>0.05) and, therefore, there are no statistical differences. Thus, it is possible to assume that groups have the same distribution regarding "',var2type,'" variable.',sep=""),paste(' (p<0.05). This means that null hypothesis is rejected (p<0.05) and, therefore, there are statistical differences. Thus, it is possible to assume that groups do not have the same distribution regarding "',var2type,'" variable.',sep="")),sep = "") 
        officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }  
    }else{
      return({0})
    }
  }
  
}

numeric.integer.numeric.integer <- function(vars.val.pass,var1type,var2type){
  
  vars.val.pass <- vars.val.pass[stats::complete.cases(vars.val.pass[,c(var1type,var2type)]),c(var1type,var2type)]
  
  #check normality for quantitative variable
  #with Kolmogorov-Smirnov (lilliefors)
  if(length(unique(stats::na.omit(vars.val.pass[,var1type])))<2 || length(stats::na.omit(vars.val.pass[,var1type]))<5){
    return({0})
  }else{
    p.value.v1 <-nortest::lillie.test(stats::na.omit(vars.val.pass[,var1type]))$p.value###The K-S test is for a continuous distribution and so MYDATA
  }
  #check normality for quantitative variable
  #with Kolmogorov-Smirnov (lilliefors)
  if(length(unique(stats::na.omit(vars.val.pass[,var2type])))<2 || length(stats::na.omit(vars.val.pass[,var2type]))<5){
    return({0})
  }else{
    p.value.v2 <-nortest::lillie.test(stats::na.omit(vars.val.pass[,var2type]))$p.value###The K-S test is for a continuous distribution and so MYDATA
  }
  
  
  #check if the variables are continuous or discrete
  if((class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "double") &&
     (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "double")){
    
    
    if(p.value.v1>0.05 && p.value.v2>0.05){
      
      #Pearson Correlation
      test <- broom::tidy(my.pearson <- stats::cor.test(vars.val.pass[,var1type],vars.val.pass[,var2type],method = "pearson"))
      test$call <- paste(var1type,"and",var2type,")",sep = " ")  
      val.test.inf <- flextable::flextable(test)
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("To analyze two numerical variables, correlations analysis is applied. Since ",var1type ,'" and "',var2type,'" have normal distribution, Pearson correlation is more appropriate.',sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("Pearson correlation presents a p-value of ",round(my.pearson$p.value,3),". This means ",ifelse(my.pearson$p.value<0.05,"a significant correlation between both variables exists."," that the independence of both variables exists."),". The value of the correlation is ",round(my.pearson$estimate,3),". Since the correlation result is ",ifelse(my.pearson$estimate>=0,"positive","negative"),", both variables are ",ifelse(my.pearson$estimate>=0,"positively","negatively")," correlated, i.e., when ",var1type," increases, ",var2type," ",ifelse(my.pearson$estimate>=0,"increases.","decreases."),sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }else{
      #Spearman Correlation
      test <- broom::tidy(my.Spearman <- stats::cor.test(vars.val.pass[,var1type],vars.val.pass[,var2type],method = "spearman"))
      test$call <- paste(var1type,"and",var2type,sep = " ")  
      val.test.inf <- flextable::flextable(test)
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("To analyze two numerical variables, correlations analysis is applied. Since ",'"',var1type ,'" and/or "',var2type,'" do not have normal distribution, Spearman correlation is more appropriate.',sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("Spearman correlation presents a p-value of ",round(my.Spearman$p.value,3),". This means that ",ifelse(my.Spearman$p.value<0.05,"a significant correlation between both variables exists.","the variables are independent.")," The value of the correlation is ",round(my.Spearman$estimate,3),". Since the correlation result is ",ifelse(my.Spearman$estimate>=0,"positive","negative"),", both variables are ",ifelse(my.Spearman$estimate>=0,"positively","negatively")," correlated, i.e., when ",var1type," increases, ",var2type," ",ifelse(my.Spearman$estimate>=0,"increases","decreases"),".",sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
  }else if(class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var2type]) == "integer"){
    
    #Spearman Correlation (one integer variable)
    test <- broom::tidy(my.Spearman <- stats::cor.test(vars.val.pass[,var1type],vars.val.pass[,var2type],method = "spearman"))
    test$call <- paste(var1type,"and",var2type,sep = " ")  
    val.test.inf <- flextable::flextable(test)
    flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.test.inf, size = 6, part = "all")), pos="after",split = TRUE) 
    officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    
    aux.text.inf <- paste('To analyze two numerical variables (in this case, one of them is an integer), correlations analysis is applied. Since we have an integer variable, Spearman correlation is more appropriate.',sep="") 
    officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
    officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    
    aux.text.inf <- paste("Spearman correlation presents a p-value of ",round(my.Spearman$p.value,3),". This means that ",ifelse(my.Spearman$p.value<0.05,"a significant correlation between both variables exists."," the independence of both variables exists.")," The value of the correlation is ",round(my.Spearman$estimate,3),". Since the correlation result is ",ifelse(my.Spearman$estimate>=0,"positive","negative"),", both variables are ",ifelse(my.Spearman$estimate>=0,"positively","negatively")," correlated, i.e., when ",var1type," increases, ",var2type," ",ifelse(my.Spearman$estimate>=0,"increases","decreases"),".",sep="") 
    officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") # blank paragraph
    officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    
  }else{
    return({0})
  }
}


optimization.cycles.1 <- function(vars.val.pass,var1type,var2type){
  
  
  while(1){
    
    if(class(vars.val.pass[,var1type])=="factor" && class(vars.val.pass[,var2type])=="factor"){
      factor.factor(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var1type])== "factor" && (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      factor.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var2type])== "factor" && (class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double")){
      numeric.integer.factor(vars.val.pass,var1type,var2type)
      break
    }
    
    if((class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double") &&
       (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      numeric.integer.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
  }
  
}


optimization.cycles.2 <- function(vars.val.pass,var1type,var2type){
  
  
  while(1){
    
    
    if(class(vars.val.pass[,var1type])== "factor" && (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      factor.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var2type])== "factor" && (class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double")){
      numeric.integer.factor(vars.val.pass,var1type,var2type)
      break
    }
    
    if((class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double") &&
       (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      numeric.integer.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var1type])=="factor" && class(vars.val.pass[,var2type])=="factor"){
      factor.factor(vars.val.pass,var1type,var2type)
      break
    }
    
    
  }
  
}


optimization.cycles.3 <- function(vars.val.pass,var1type,var2type){
  
  
  while(1){
    
    
    if((class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double") &&
       (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      numeric.integer.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var1type])=="factor" && class(vars.val.pass[,var2type])=="factor"){
      factor.factor(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var1type])== "factor" && (class(vars.val.pass[,var2type]) == "numeric"|| class(vars.val.pass[,var2type]) == "integer" || class(vars.val.pass[,var2type]) == "double")){
      factor.numeric.integer(vars.val.pass,var1type,var2type)
      break
    }
    
    if(class(vars.val.pass[,var2type])== "factor" && (class(vars.val.pass[,var1type]) == "numeric"|| class(vars.val.pass[,var1type]) == "integer" || class(vars.val.pass[,var1type]) == "double")){
      numeric.integer.factor(vars.val.pass,var1type,var2type)
      break
    }
    
  }
  
}


inf.writing <- function(what="inference",vars.name.pass=NULL,vars.val.pass=NULL,analysis=NULL){
  
  if(what=="inference"){
    
    if(analysis=="header1"){
      
      val.header1 <-  paste("Inference Analysis for All Variables Combination in Dataset:", sep = "")
      officer::body_add_par(my_doc, val.header1, style = "heading 1",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste('An assessment of the normality of data is a prerequisite for many statistical tests because normal data is an underlying assumption in parametric tests. The variables used to test the normality need to be numerical. Additionally, to numerically test the normality, the length of the sample should be taken into account. If the length is smaller than 50 records we use "Shapiro-Wilk test" or else we use "Kolmogorov-Smirnov test".',sep="")
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("Thus, normality test was done for the numerical variables. The results of the tests used -- Shapiro-Wilk or Kolmogorov-Smirnov test - allows us concluding that, with a 95% of confidence, the null hypothesis is rejected if p<0.05/non-rejected if p>0.05, i.e., there is evidence/there is no evidence to reject the null hypothesis and it may not be/ be considered the existence of normality, respectively.",sep="")
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("In this sense, the numeric variable has not/has normal distribution and, consequently, non-parametric/parametric tests should be used with that variable.",sep="")
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("To compare two or more categorical variables, a cross-tabulation (also called the contingency table) is the most adequate option. However, to analyze the statistical differences, the chi-squared test or fisher test, for independence, should be applied to the crosstab.",sep="")
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      aux.text.inf <- paste("To analyze two numerical variables, correlations analysis is applied. If variables have/have not normal distribution, Pearson / Spearman correlation is more appropriate, respectively.",sep="") 
      officer::body_add_par(my_doc, aux.text.inf, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    } 
    
    if(analysis=="header2"){
      
      val.header2 <-  paste('Inference Analysis between Variable "', vars.name.pass[1], '" and "',vars.name.pass[2],'":', sep = "")
      officer::body_add_par(my_doc, val.header2, style = "heading 2",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="res"){
      
      var1type <- vars.name.pass[1]
      var2type <- vars.name.pass[2]
      
      #Optimization considering variable type to increase efficiency
      nfactor.count <- length(which(sapply(vars.val.pass, class)=="factor"))
      nnumeric.count <- length(which(sapply(vars.val.pass, class) %in% c("numeric","integer","double")))
      
      if(nfactor.count > nnumeric.count)
      {
        optimization.cycles.1(vars.val.pass,var1type,var2type)
      }else if(nfactor.count == nnumeric.count){
        optimization.cycles.2(vars.val.pass,var1type,var2type)
      }else{
        optimization.cycles.3(vars.val.pass,var1type,var2type)
      }
      
      
    }
  }
}


file1 <- file_in

if(is.null(file1)){return(NULL)}

  a<-readLines(path,n=2)
  b<- nchar(gsub(";", "", a[1]))
  c<- nchar(gsub(",", "", a[1]))
  d<- nchar(gsub("\t", "", a[1]))
  if(b<c&b<d){
    #r <- utils::read.table(path,sep=";",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep=";",stringsAsFactors = FALSE,nrows=1)
    
    SEP=";"
  }else if(c<b&c<d){
    #r <- utils::read.table(path,sep=",",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep=",",stringsAsFactors = FALSE,nrows=1)
    SEP=","
  }else{
    #r <- utils::read.table(path,sep="\t",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep="\t",stringsAsFactors = FALSE,nrows=1)
    SEP="\t"
  }
  
  
  #r1<-utils::read.table(path,sep=SEP,stringsAsFactors = FALSE,nrows=1,skip = 1,na.strings = c("NA",""))
  r1<-utils::read.table(path,sep=SEP,stringsAsFactors = FALSE,nrows=1,skip = 1)
  aciertos <- 0
  for(i in 1: ncol(r)){
    if(class(r[1,i])==class(r1[1,i])){
      aciertos <- aciertos+1
    }
  }
  
  HEADER=T
  if(aciertos==ncol(r)){
    HEADER=F
  }
  
data <- utils::read.table(path,sep=SEP,header=HEADER)

#Inference Analysis
  
  #variables combinations
  data.vars.vector <- names(data)
  
  #type of selected vars
  if(var.type!="all"){
      data.vars.vector <- data.vars.vector[which(sapply(data,class)==var.type)]

      if (length(data.vars.vector) < 2){
	 warning(paste("Not enough variables of type", var.type, "to proceed with Inference Reporting. Suggestion: use var.type input equal to all", sep=" "))
	 return()
      }

  }
  
  var.comb <- t(utils::combn(data.vars.vector,2))
  
  inf.writing(what = "inference",vars.name.pass=NULL,vars.val.pass=NULL,analysis="header1")
   
  tictoc::tic("Analysis")
    for(comb in 1:nrow(var.comb)){
      inf.writing(what = "inference",vars.name.pass=var.comb[comb,],vars.val.pass=data[,var.comb[comb,]],analysis="header2")
      aux.res.inf <- inf.writing(what = "inference",vars.name.pass=var.comb[comb,],vars.val.pass=data[,var.comb[comb,]],analysis="res")
      if(class(aux.res.inf)=="numeric"){
        if(aux.res.inf==0){
          officer::body_add_par(my_doc, "Analysis not possible due to data constraints!", style = "Normal",pos="after") 
          officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        }
      }
    }
    
  save.image(file = paste(aux.file.rdata))
  warning(Sys.time(),"Passed Inference\n\n")
  tictoc::toc()

tictoc::tic("Writing DOC")
print(my_doc, target = paste(aux.file.path))
tictoc::toc()


save.image(file = paste(aux.file.rdata))
warning(Sys.time(),"Passed Final Print\n\n")
tictoc::toc()

if(grepl("lin",tolower(my.OS))){
  close(aux.file.report)
}
warning("Passed Final Print")

sink(type = "message")
sink()
#end and print
}

#' Inference Report in .docx file
#'
#' This R Package asks for a .csv file with data and returns a report (.docx) with Inference Report concerning all possible combinations of variables (i.e. columns).
#' 
#' @import officer nortest flextable broom tictoc stats utils 
#' @importFrom Rdpack reprompt
#'
#' @param path (Optional) A character vector with the path to data file. If empty character string (""), interface will appear to choose file. 
#' @param var.type (Optional) The type of variables to perform analysis, with possible values: "all", "numeric", "integer", "double", "factor", "character". 
#' @return The output
#'   will be a document in the temp folder (tempdir()).
#' @examples
#' \donttest{
#' library(docinfeR)
#' data(iris)
#' dir = tempdir()
#' write.csv(iris,file=paste(dir,"iriscsvfile.csv",sep=""))
#' docinfeR(path=paste(dir,"iriscsvfile.csv",sep=""))
#' }
#' @references
#' 	\insertAllCited{}
#' @export
docinfeR <- function(path="", var.type="all"){
	main(path,var.type=var.type)
}
