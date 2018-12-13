
#First of all you set up the current directory where the files are.
# it's not longer necessary to set working directory.
files=dir('lectura/movimientos/', pattern = '\\.xls$')

##two empty vector are create to save the files.
##the files length relies on the files number the folder actually has and
##the files that match with \\.xls$

carga=vector('list',length(files)) #This is a string that have the files's names, the
leido=vector('list', length(files))

library(XLConnect) # load XLconnect library that allows to read every file

#First round with a FOR-LOOP for loading the files names; i.e c("files002.xls"...."file0101.xls...xxxxx.xls")
for (i in seq_along(files)){
   carga[[i]]=loadWorkbook(paste0('lectura/movimientos/',files[i]), create = FALSE)

}
# The names are already loaded as welll the workbooks ,now we need to convert the file as xls object into R Enviorenment
##vector que leer con REadWorkSHEEt
##In this case the sheet I read is 'Detalle_camion _Pala'

for (tt in seq_along(carga)){
  leido[[tt]]=readWorksheet(carga[[tt]],sheet = 'Detalle_Camion_Pala')
  
}   



## The first clean up task dependent on what you need to do,rmacion
#save the original ..leido..
leido2=leido
#create a empty vector to be the warehouse of next files.
leido3=vector('list', length(leido2))
#So far, my files have NA's, garbage I dont really need, point.So that I do filter over certain conditions
# 
for(p in seq_along(leido2)){
  for( o in seq_along(leido2[[p]])){
    if (!is.logical(leido2[[p]][[o]])){
      leido3[[p]][[o]]=leido2[[p]][[o]]
    
    }
  }
}
## creating columns
##I know the file format and arquitecture, where the variables are and the column belongings
#So, I do some task for retrieve info from the column
library(stringr)


for (j in seq_along(leido3)) {
   if (any(str_detect(pattern="Dia", leido3[[j]]))){ ##if 'Dia' belongs to this file go forward
      for (k in (1:length(leido3[[j]][[26]]))){
        leido3[[j]][[14]][k]=leido3[[j]][[8]][8]##save the prior element into a empty column into the same file.
      } 
     
   }else if (any(str_detect(pattern="Noche", leido3[[j]]))){ #if 'Noche' belongs to this file go forward
     for (l in (1:length(leido3[[j]][[26]]))){
       leido3[[j]][[14]][l]=leido3[[j]][[8]][8]##save the prior element into a empty column into the same file.
     }
  
   }
}

## a empty dataframe, the files are still elements list, I needed to turn into  dataframe to go on cleaning up
dataframe=vector('list', length(leido3))
# Selecting the column you need to retrieve from the excel file.
for( g in seq_along(leido3)){
  
  dataframe[[g]]=data.frame(leido3[[g]][[1]],leido3[[g]][[9]],leido3[[g]][[11]],leido3[[g]][[14]],
                          leido3[[g]][[16]],leido3[[g]][[20]],leido3[[g]][[22]],leido3[[g]][[23]],
                          leido3[[g]][[24]],leido3[[g]][[25]],leido3[[g]][[26]] )
}
 ### intento de pegado de todos los 772 dataframes y un solo archivo
#turning into a dataframe objetc that contains all file into one.
todo=data.frame(stringsAsFactors = FALSE)
for (h in seq_along(dataframe)){
  todo=rbind(todo,dataframe[[h]])
}



#save again the prior element 'todo' creating a file 'Todo2'
todo2=todo
###El criterio aqui es eliminar los NA, parece que el algoritmo de is.na no elimina los NA q
## que esta dentro del dataframe
## Eliminating the Na's from every row.
missing=which(is.na(todo2))
todo3=todo2[-missing,]
##los nombres
nombres=c('Truck','Load Location','Dump Location','Fecha','Shovel','Material Type','Loads','Dumps',
          'Cycle Distance (Kilometros)','Cycle (mins)','Toneladas')
##nombra las columnas del data frame
colnames(todo3)=nombres
# we need to detect where the files start and where end.
# Remebember I come together every file without rule out the head.
head_names=which(todo3[1]=='Truck')


##subset taking out the head_names
todo4=todo3[-head_names, ]
todo5=todo4
##this is option for you, because at this point we got the dataframe unificated
##from 772 excel files to 1 excel file.
write.csv(todo5, 'C:/Users/DANIEL/documents/projects/orange/lectura_archivos.csv')


 