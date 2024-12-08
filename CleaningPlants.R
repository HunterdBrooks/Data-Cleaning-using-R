

library (readxl)
library(dplyr)
library(stringr)
library(openxlsx)
setwd("/Users/hunterbrooks/Desktop/InvasivePlants")
getwd()

#Sheet 1
InvasivesPoint0=read_excel("Previous_Invasive_Data_(Prior_to_2024).xlsx",sheet="Previous_Invasives_Point_0")

NAsheet1=InvasivesPoint0%>%
  filter(is.na(Species) & is.na(`Comments/Description`))

Sheet1=InvasivesPoint0%>%
  rename(Name='Comments/Description')%>%
  #Removing NA's for both name and species
  filter(!(is.na(Species) & is.na(Name)))%>%
  mutate(Species = str_replace_all(Species, regex("sweet\\s*gum", ignore_case = TRUE), "Sweetgum"))%>%
  mutate(Species = str_replace_all(Species, regex("vernal\\s*pool", ignore_case = TRUE), "Vernal Pool"))
  
Invasive<-c("privet", "holly", "japanese","chinese","rose", "olive", "berry","heaven","Mimosa","Callery","Princess","ash","HEHE","MIVI","LOJA","Bamboo","Ivy","COCO3","Mahonia","ELUM","WISI","LEBI2","Microstegium","bittersweet","fescue","ALJU","honeysuckle","kudzu","Mugwort","romu","Wisteria","sweet", "golden", "alligator", "creeper", "European","Ligustrum", "Eleagnus", "Lonicera", "pear", "Trifoliolate","Stiltgrass", "westeria", "PATO2", "HOCO3", "VETH", "AMBR7", "LECU", "VIMA", "lepadeza", "Arbor","Melia","racer","tbd", "Elaeagnus", "umbrella","tunnel", "glory", "test", "application", "ncma", "open", "seedlings","applications", "pruning","trees", "sad", "diameter", "nice", "big", "Signage", "wetland","huge","Autumn","pennywort", "resilience","atv", "Asclepius","Nandina","Adders", "AIAL", "Ailanthus","Albizia","Aranaria","Artemisia", "Bear","BEBE", "Bull Thistle","Cane","Dogbane","boxwood","Creeping jenny", "ELPU2", "lvy","EUFO5","buttercup","Firmiana","Glechoma","Hedera","Hop","Humulus","Hedera","cornuta", "Invasive","Johnson Grass","lespedeza","lezpedeza","LILU2","LISI","Mondo grass","Morus alba", "Musk thistle","NAD0","Parrots feather","Pueraria montana","Pumo","Pyrus","Rosa multiflora","lespedeza", "Stilt", "Vinca major","Berry","ivy","kudzu","olive","privet","Wisteria", "Chinese","Honeysuckle","bamboo","Johnsongrass","Liriope spicata", "PYCA80", "Red tipped photina","olive","MIVI","HEHE","Hedera","privet","wisteria","Japanese","pear","LEBI2","LECU","LOJA","kudzu","berry","bamboo","emerald", "virginsbower", "mimosa", "LOJO", "Stiltgrass","Heaven","PEFR4","ELUM", "BEBE", "COCO3","Lespedeza","lezpedeza", "application", "Pueraria", "Firmiana", "Alternathera", "Pueria", "japonicus","japonicum", "Scotch", "Clematis", "Tatarian", "Multiflora", "Wintercreeper", "Misi", "PHAO8", "bicolor","westeria","Elaegnus","Stilitgrass", "Stilt", "Celastrus", "brevipedunculata","Microstegium", "Liriope", "Ampelopsis", "English","Ligustrum","Ailanthus", "Hop", "Umbellata", "Podophyllum","Buttercup","Asiatic", "Ash")

NativesSheet1Cleaned<-Sheet1%>%
  filter(!grepl(paste(Invasive, collapse="|"), Species, ignore.case = TRUE))%>%
  filter(!grepl(paste(Invasive, collapse="|"), Name, ignore.case = TRUE))

InvasivesSheet1Cleaned <- Sheet1 %>%
  filter(grepl(paste(Invasive, collapse = "|"), Species, ignore.case = TRUE) | 
grepl(paste(Invasive, collapse = "|"), Name, ignore.case = TRUE))

#x<-table(Sheet1Cleaned$Species)
#xFiltered <- x[x>15]
#print(xFiltered)

#y<-table(Sheet1Cleaned$Name)
#yFiltered <- y[y>15]
#print(yFiltered)

NativesPoint0<-NativesSheet1Cleaned%>%
  rename('Comments/Description'=Name)

InvasivesPoint0<-InvasivesSheet1Cleaned%>%
  rename('Comments/Description'=Name)

#Sheet 2 
InvasivesLine1=read_excel("Previous_Invasive_Data_(Prior_to_2024).xlsx",sheet="Previous_Invasives_Line_1")

NAsheet2=InvasivesLine1%>%
  filter((is.na(Species) & is.na(`Comments/Description`)))

Sheet2=InvasivesLine1%>%
  rename(Name='Comments/Description')%>%
  #Removing NA's for both name and species
  filter(!(is.na(Species) & is.na(Name)))
  
NativesSheet2Cleaned<-Sheet2%>%
  filter(!grepl(paste(Invasive, collapse="|"), Species, ignore.case = TRUE))%>%
  filter(!grepl(paste(Invasive, collapse="|"), Name, ignore.case = TRUE))

InvasivesSheet2Cleaned <- Sheet2 %>%
  filter(grepl(paste(Invasive, collapse = "|"), Species, ignore.case = TRUE) | 
           grepl(paste(Invasive, collapse = "|"), Name, ignore.case = TRUE))


NativesLine1<-NativesSheet2Cleaned%>%
  rename('Comments/Description'=Name)
InvasivesLine1<-InvasivesSheet2Cleaned%>%
  rename('Comments/Description'=Name)

#Sheet 3
InvasivesPolygon2=read_excel("Previous_Invasive_Data_(Prior_to_2024).xlsx",sheet ="Previous_Invasives_Polygon_2")%>%
  select(OBJECTID,Species,`Comments/Description`)

NAsheet3=InvasivesPolygon2%>%
  filter((is.na(Species) & is.na(`Comments/Description`)))

Sheet3=InvasivesPolygon2%>%
  rename(Name='Comments/Description')%>%
  filter(!(is.na(Species) & is.na(Name)))
  
NativesSheet3Cleaned<-Sheet3%>%
  filter(!grepl(paste(Invasive, collapse="|"), Species, ignore.case = TRUE))%>%
  filter(!grepl(paste(Invasive, collapse="|"), Name, ignore.case = TRUE))

InvasivesSheet3Cleaned <- Sheet3 %>%
  filter(grepl(paste(Invasive, collapse = "|"), Species, ignore.case = TRUE) |
           grepl(paste(Invasive, collapse = "|"), Name, ignore.case = TRUE))


NativesPolygon2=NativesSheet3Cleaned%>%
  rename('Comments/Description'=Name)
InvasivesPolygon2=InvasivesSheet3Cleaned%>%
  rename('Comments/Description'=Name)



NativeData <- list('NativesPoint0' = NativesPoint0,'NativesLine1'=NativesLine1, 'NativesPolygon2' = NativesPolygon2)

InvasiveData = list('InvasivesPoint0'=InvasivesPoint0, 'InvasivesLine1'=InvasivesLine1, 'InvasivesPolygon2'=InvasivesPolygon2)
  
NAdatasetsNative<-list('NAsForPoint0'=NAsheet1,'NAsForLine1'=NAsheet2, 'NAsForPolygon2'=NAsheet3)


write.xlsx(NativeData, file='NativesData.xlsx')
write.xlsx(InvasiveData, file='InvasivesData.xlsx')
write.xlsx(NAdatasetsNative, file='NAdataNatives.xlsx')



