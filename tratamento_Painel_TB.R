################################################################################################################################
################################################       15 DE SETEMBRO DE 2025        ###########################################
###############################################       PAINEL TB, INDICADORES TB e AVALIAÇÃO DE DESEMPENHO######################
##############################################            versão R 4.0.3            ############################################
############################################            CASSIA CP MENDICINO         ############################################
###########################################                  ONCE AGAIN            ############################################ 
############################################MY LORDY GIVE-ME STRAING TO CARRY ON    ############################################
################################## SANCTA MARIA, MATER DEI, ORA PRO NOBIS PECCATORIBUS #########################################
################################################################################################################################


#FONTE DE INFORMAÇÕES: 

#RELATORIO SISTEMA DE NOTIFICAÇÃO DE AGRAVOS DE NOTIFICAÇÕES PARA TUBERCULOSE DE MINAS GERAIS A PARTIR DO ANO DE 2000
#TABELAS DE MUNICÍPIOS PDR 2024 COM POPULAÇÃO


rm(list = ls())
library(writexl)#Exportar em Excel
library(data.table)
library(tidyr)# função pivot (emparelhamento de colunas)
library(dplyr) 
library(openxlsx)
library(readxl)
library(foreign)
library(lubridate)
library(dplyr) #função mutate
library(readr)#EXPORTAR EM CSV


#### IMPORTAÇÃO 
#HOME
TB_bruto<-read.dbf("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/TUBENET.dbf") #124.033
MUNICIPIOS<-read.xlsx("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/PDR2026_comPOP.xlsx") 

#SES
TB_bruto<-read.dbf("C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/TUBENET.dbf") #125.644
MUNICIPIOS<-read.xlsx("C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/PDR2026_comPOP.xlsx") 


##################################################  TRATAMENTO GERAL DAS BASES #####################################################

## BASE TB
###   SELEÇÃO DAS VARIÁVEIS
TB<-TB_bruto[,c("NU_NOTIFIC","ID_MN_RESI","DT_DIAG","SITUA_ENCE","TRATAMENTO","FORMA","DT_NASC","NU_IDADE_N","CS_SEXO",
                "CS_GESTANT","CS_RACA","CS_ESCOL_N","POP_LIBER","POP_RUA","POP_IMIG","POP_SAUDE","NU_CONTATO","NU_COMU_EX",
                "BACILOSC_E","BACILOSC_2","BACILOS_E2","CULTURA_ES","CULTURA_OU","TEST_MOLEC","AGRAVAIDS","HIV","EXTRAPU1_N",
                "EXTRAPU2_N","EXTRAPUL_O","AGRAVALCOO","AGRAVDIABE","AGRAVDOENC","AGRAVOUTRA","AGRAVOUTDE","AGRAVDROGA","AGRAVTABAC")]

##  Considerar apenas os residentes de Minas Gerais  ##
TB <- subset(TB, grepl("^31", as.factor(ID_MN_RESI))) 

##  Excluir mudanças de diagnóstico  ##
TB$SITUA_ENCE <- gsub("06","6",TB$SITUA_ENCE)
TB <- subset(TB, SITUA_ENCE != "6" | is.na(SITUA_ENCE))

##  Considerar a partir de 2010  ##  
TB$Ano<-format(TB$DT_DIAG,"%Y")
TB$Ano<-as.integer(TB$Ano)
TB<-subset(TB,Ano>=2010)

##  Excluir ID_MN_RESI inválido: 1 ocorrência
TB <- subset(TB, ID_MN_RESI != "310000" | is.na(ID_MN_RESI))

## Tratamento das variáveis:
TB<-TB %>% 
  mutate(SITUA_ENCE = case_when(
    SITUA_ENCE %in% c("1")~ "Cura",
    SITUA_ENCE %in% c("01")~ "Cura",
    SITUA_ENCE %in% c("2","10")~ "Abandono", #Juntar abandono primário com abandono#
    SITUA_ENCE %in% c("3","4")~ "Óbito", #juntar com óbitos por outras causas
    SITUA_ENCE %in% c("5")~ "Transferência",
    SITUA_ENCE %in% c("7")~ "TB-DR",
    SITUA_ENCE %in% c("8")~ "Mudança de Esquema",
    SITUA_ENCE %in% c("9")~ "Falência",
    TRUE ~ "Caso não encerrado"))

TB<-TB %>% 
  mutate(TRATAMENTO = case_when(
    TRATAMENTO %in% c("1")~ "Caso novo",
    TRATAMENTO %in% c("2")~ "Recidiva",
    TRATAMENTO %in% c("3")~ "Reingresso após abandono",
    TRATAMENTO %in% c("4")~ "Não sabe",
    TRATAMENTO %in% c("5")~ "Transferência",
    
    
    
    TRATAMENTO%in% c("6")~ "Pós-óbito",
    TRUE ~ "Não sabe"))

TB<-TB %>% 
  mutate(FORMA = case_when(
    FORMA %in% c("1")~ "Pulmonar",
    FORMA %in% c("2")~ "Extrapulmonar",
    FORMA %in% c("3")~ "Pulmonar e extrapulmonar",
    TRUE ~ "Dado não informado"))

#Criação da variável PVHA a partir das informações das colunas AGRAVAIDS E HIV:
TB <- TB%>%
  mutate(
   PVHA= case_when(
     AGRAVAIDS == "1" | HIV == "1" ~ "Coinfecção TB/HIV",
     AGRAVAIDS == "2" | HIV == "2" ~ "Infecção TB",
    TRUE ~ "Ignorado"
    )
  )

## BASE PDR/2026
names(MUNICIPIOS)[names(MUNICIPIOS)=="CÓDIGO"]<- "ID_MN_RESI"
names(MUNICIPIOS)[names(MUNICIPIOS)=="COD..MICRO/REGIÃO"]<- "COD..MICRO"




###################################################  PREVALENCIAS E TOTAL DE CASOS  #####################################################


## BANCO MUNICÍPIOS  ##

MUNICIPIOS1<-MUNICIPIOS[,c("ID_MN_RESI","NM_MUNICIP","REGIÃO.DE.SAÚDE","Unidade.Regional.de.Saúde",
                          "REGIÃO.AMPLIADA.DE.SAÚDE","POP_Municipio","POP_URS")]
#Populações macro e micro
MUNICIPIOS1 <- MUNICIPIOS1 %>%
  group_by(REGIÃO.DE.SAÚDE) %>%
  mutate(POP_micro = sum(POP_Municipio, na.rm = TRUE)) %>%
  ungroup()
MUNICIPIOS1 <- MUNICIPIOS1 %>%
  group_by(REGIÃO.AMPLIADA.DE.SAÚDE) %>%
  mutate(POP_macro = sum(POP_Municipio, na.rm = TRUE)) %>%
  ungroup()

#Alterar o nome das MICRO, regional e de municípios para localização nos mapas:
MUNICIPIOS1$NM_MUNICIP<-gsub("Gouvêa", "Gouveia",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("Queluzita", "Queluzito",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("São Thomé das Letras", "São Tomé das Letras",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("Brasópolis", "Brazópolis",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("Passa-Vinte", "Passa Vinte",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("Olhos-D'água", "Olhos-d'Água",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("Pingo-D'água", "Pingo-d'Água",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$NM_MUNICIP<-gsub("São João Del Rei", "São João del Rei",MUNICIPIOS1$NM_MUNICIP)
MUNICIPIOS1$Unidade.Regional.de.Saúde<-gsub("GOV. VALADARES", "GOVERNADOR VALADARES",MUNICIPIOS1$Unidade.Regional.de.Saúde )
MUNICIPIOS1$Unidade.Regional.de.Saúde<-gsub("CEL. FABRICIANO", "CORONEL FABRICIANO",MUNICIPIOS1$Unidade.Regional.de.Saúde )
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("FRUTAL / ITURAMA", "FRUTAL/ITURAMA",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("UBERLÂNDIA / ARAGUARI", "UBERLÂNDIA/ARAGUARI",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("PATROCÍNIO / MONTE CARMELO", "PATROCÍNIO/MONTE CARMELO",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("ALFENAS / MACHADO", "ALFENAS/MACHADO",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("JANAÚBA / MONTE AZUL", "JANAÚBA/MONTE AZUL",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("ALMENARA/ JACINTO", "ALMENARA/JACINTO",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("MINAS NOVAS / TURMALINA / CAPELINHA", "TURMALINA/MINAS NOVAS/CAPELINHA",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub(" TAIOBEIRAS", "TAIOBEIRAS",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("PEÇANHA/SÃO JOÃO EVANGELISTA/SANTA MARIA DO SUAÇUI", "PEÇANHA/SÃO JOÃO EVANGELISTA/SANTA MARIA DO SUAÇUÍ",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("TEÓFILO OTONI / MALACACHETA", "TEÓFILO OTONI/MALACACHETA",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("CORONEL FABRICIANO / TIMÓTEO", "CORONEL FABRICIANO/TIMÓTEO",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("BELO HORIZONTE / NOVA LIMA / SANTA LUZIA", "BELO HORIZONTE/NOVA LIMA/SANTA LUZIA",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("LAGOA DA PRATA / SANTO ANTÔNIO DO MONTE", "LAGOA DA PRATA/SANTO ANTÔNIO DO MONTE",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub(" PIUMHI", "PIUMHI",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub(" CAMPO BELO", "CAMPO BELO",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub(" LIMA DUARTE", "LIMA DUARTE",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("SÃO JOÃO NEPOMUCENO / BICAS", "SÃO JOÃO NEPOMUCENO/BICAS",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("LEOPOLDINA / CATAGUASES", "LEOPOLDINA/CATAGUASES",MUNICIPIOS1$REGIÃO.DE.SAÚDE)
MUNICIPIOS1$REGIÃO.DE.SAÚDE<-gsub("ALÉM  PARAIBA", "ALÉM PARAÍBA",MUNICIPIOS1$REGIÃO.DE.SAÚDE)


#Código pra abrir arquivo JSON e conferir os nomes:
#library(jsonlite)
#dados <- fromJSON("C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Geoprocessamento/Json/MG_Municipios_2022.json", flatten = TRUE)
#str(dados)
#names(dados)
#str(dados, max.level = 4)
#municipiosJSON <- data.frame(
  #NM_MUN = dados$objects$MG_Municipios_2022$geometries$properties.NM_MUN,
  #CD_MUN = dados$objects$MG_Municipios_2022$geometries$properties.CD_MUN
#)

#Conferir os nomes dos municípios:
#names(municipiosJSON)[names(municipiosJSON)=="NM_MUN"]<- "NM_MUNICIP"
#Conferencia<-merge(MUNICIPIOS1,municipiosJSON, by = "NM_MUNICIP", all = TRUE)
#OBSERVAÇÃO
#Santana do Gambaréu possui 3 notificações, mas o arquivo JSON não localizou e o mapa por município fico sem caso (??)


## BANCO TB  ##

TB1<-TB[,c("NU_NOTIFIC","ID_MN_RESI","Ano","TRATAMENTO","FORMA","SITUA_ENCE")] 

# LINKAR BANCO TB E MUNICIPIOS
## VARIÁVEL "Unidade.Regional.de.Saúde"
TB1 <- merge(MUNICIPIOS1, TB1, by = "ID_MN_RESI", all = TRUE)

#Municip_sem_TB<-subset(TB1,NU_NOTIFIC=="" | is.na (NU_NOTIFIC)) 
#8 municípios em nenhuma notificação
#Antônio Prado de Minas
#Carvalhópolis
#Conceição das Pedras
#Ibituruna
#Lamim
#Queluzito
#Santo Antônio do Rio Abaixo
#Senhora do Porto


## EXportação o para o painel TB
#write.csv(TB1,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Prevalencias_casos.csv") 
#write.csv(TB1,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Prevalencias_casos.csv")



###################################################  PERFIL DEMOGRÁFICO #####################################################

###   SELEÇÃO DAS VARIÁVEIS",
TB2<-TB[,c("NU_NOTIFIC","ID_MN_RESI","Ano","DT_DIAG","DT_NASC","NU_IDADE_N","TRATAMENTO","FORMA","SITUA_ENCE",
            "CS_SEXO","CS_ESCOL_N","CS_GESTANT","CS_RACA","POP_LIBER","POP_RUA","POP_IMIG","POP_SAUDE", "PVHA",
           "EXTRAPU1_N","EXTRAPU2_N","EXTRAPUL_O","AGRAVALCOO","AGRAVDIABE","AGRAVDOENC","AGRAVOUTDE",
           "AGRAVDROGA","AGRAVTABAC")]
                
MUNICIPIOS2<-MUNICIPIOS[,c("NM_MUNICIP","ID_MN_RESI","REGIÃO.DE.SAÚDE","Unidade.Regional.de.Saúde",
                           "REGIÃO.AMPLIADA.DE.SAÚDE")]

## linkagem das bases:
TB2 <- merge(MUNICIPIOS2, TB2, by = "ID_MN_RESI", all = TRUE)
#Municip_sem_TB<-subset(TB2,NU_NOTIFIC=="" | is.na (NU_NOTIFIC)) #10 municípios em nenhuma notificação
TB2<-TB2 %>% select(-ID_MN_RESI)

#TRATAMENTO DE VARIÁVEIS
#IDADE NA DATA DO DIAGNÓSTICO
#Variável IDADE e IDADE_cat;
TB2$IDADE<-round((as.numeric(TB2$DT_DIAG-TB2$DT_NASC))/365.25)
#Quando não tiver data de nascimento usar a variável NU_IDADE_N para preencher os dados da variável Idade:
TB2$idade_notif<- as.numeric(TB2$NU_IDADE_N - 4000)
TB2<- TB2 %>% 
  mutate (IDADE = ifelse (is.na(IDADE),idade_notif, IDADE))
#EXcluir número negativo:
TB2 <- TB2 %>%
  mutate(IDADE = ifelse(IDADE < 0, NA, IDADE))
TB2<-TB2 %>% select(-DT_DIAG,-DT_NASC,-NU_IDADE_N,-idade_notif)

#IDADE_CAT: Fonte categorização: Faixa etária Sinan 
TB2$IDADEcat<-rep(NA,length(TB2$IDADE))
TB2$IDADEcat[TB2$IDADE<1]<-"Inferior a 1 ano"
TB2$IDADEcat[1<=TB2$IDADE & TB2$IDADE<5]<-"1 a 4 anos"
TB2$IDADEcat[5<=TB2$IDADE & TB2$IDADE<10]<-"5 a 9 anos"
TB2$IDADEcat[10<=TB2$IDADE & TB2$IDADE<15]<-"10 a 14 anos"
TB2$IDADEcat[15<=TB2$IDADE & TB2$IDADE<20]<-"15 a 19 anos"
TB2$IDADEcat[20<=TB2$IDADE & TB2$IDADE<35]<-"20 a 34 anos"
TB2$IDADEcat[35<=TB2$IDADE & TB2$IDADE<50]<-"35 a 49 anos"
TB2$IDADEcat[50<=TB2$IDADE & TB2$IDADE<65]<-"50 a 64 anos"
TB2$IDADEcat[65<=TB2$IDADE & TB2$IDADE<80]<-"65 a 79 anos"
TB2$IDADEcat[TB2$IDADE>=80]<-"Superior a 80 anos"
TB2$IDADEcat=as.factor(TB2$IDADEcat)

#Ordem idade: para ordenar no BI:
TB2$IDADEordem<-rep(NA,length(TB2$IDADEcat))
TB2$IDADEordem[TB2$IDADEcat=="Inferior a 1 ano"]<-10 
TB2$IDADEordem[TB2$IDADEcat=="1 a 4 anos"]<-9
TB2$IDADEordem[TB2$IDADEcat=="5 a 9 anos"]<-8
TB2$IDADEordem[TB2$IDADEcat=="10 a 14 anos"]<-7
TB2$IDADEordem[TB2$IDADEcat=="15 a 19 anos"]<-6
TB2$IDADEordem[TB2$IDADEcat=="20 a 34 anos"]<-5
TB2$IDADEordem[TB2$IDADEcat=="35 a 49 anos"]<-4
TB2$IDADEordem[TB2$IDADEcat=="50 a 64 anos"]<-3
TB2$IDADEordem[TB2$IDADEcat=="65 a 79 anos"]<-2
TB2$IDADEordem[TB2$IDADEcat=="Superior a 80 anos"]<-1
TB2$IDADEordem=as.numeric(TB2$IDADEordem)

#SEXO:
TB2<-TB2 %>% 
  mutate(SEXO = case_when(
    CS_SEXO %in% c("F")~ "Feminino",
    CS_SEXO %in% c("M")~ "Masculino",
    CS_SEXO %in% c("I")~ "Indeterminado",
    TRUE ~ "Ignorado"))
TB2<-TB2 %>% select(-CS_SEXO)

#Raça/cor
TB2<-TB2 %>% 
  mutate(Raça = case_when(
    CS_RACA %in% c("1")~ "Branca",
    CS_RACA %in% c("2")~ "Preta",
    CS_RACA %in% c("3")~ "Amarela",
    CS_RACA %in% c("4")~ "Parda",
    CS_RACA %in% c("5")~ "Indígena",
    CS_RACA %in% c("9")~ "Ignorado",
    TRUE ~ "Ignorado"))
TB2<-TB2 %>% select(-CS_RACA)

#Escolaridade:
TB2<-TB2 %>% 
  mutate(Escolaridade = case_when(
    CS_ESCOL_N %in% c("00")~ "Analfabeto",
    CS_ESCOL_N %in% c("01","02","03","04")~ "Ensino Fundamental",
    CS_ESCOL_N %in% c("05","06")~ "Ensino médio",
    CS_ESCOL_N %in% c("07","08")~ "Ensino superior",
    CS_ESCOL_N %in% c("09")~ "Ignorado",
    CS_ESCOL_N %in% c("10")~ "Não se aplica",
    TRUE ~ "Ignorado"))
TB2<-TB2 %>% select(-CS_ESCOL_N)

#Gestantes
TB2<-TB2 %>% 
  mutate(Gestante = case_when(
    CS_GESTANT %in% c("1","2","3","4")~ "Sim",
    CS_GESTANT %in% c("5")~ "Não",
    CS_GESTANT %in% c("6")~ "Não se aplica",
    CS_GESTANT %in% c("9")~ "Ignorado",
    TRUE ~ "Ignorado"))
TB2<-TB2 %>% select(-CS_GESTANT)

#Tratar a categoria "Não se aplica"
#NÃO SE APLICA: Homens, mulheres abaixo de 10 anos e  acima de 60 anos
TB2$Gestante <- ifelse(
  TB2$SEXO == "Masculino" | TB2$IDADE < 10 | TB2$IDADE > 60, 
  "Não se aplica", 
  TB2$Gestante
)

#Condição social:
TB2<-TB2 %>% 
  mutate(Privado_liberdade = case_when(
    POP_LIBER %in% c("1")~ "Sim",
    POP_LIBER %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB2<-TB2 %>% select(-POP_LIBER)
TB2<-TB2 %>% 
  mutate(Situacao_rua = case_when(
    POP_RUA %in% c("1")~ "Sim",
    POP_RUA %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB2<-TB2 %>% select(-POP_RUA)
TB2<-TB2 %>% 
  mutate(Imigrante = case_when(
    POP_IMIG %in% c("1")~ "Sim",
    POP_IMIG %in% c("2")~ "Não",
   TRUE ~ "Sem informação"))
TB2<-TB2 %>% select(-POP_IMIG)
TB2<-TB2 %>% 
  mutate(Profissional_saude = case_when(
    POP_SAUDE %in% c("1")~ "Sim",
    POP_SAUDE %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB2<-TB2 %>% select(-POP_SAUDE)
TB2<-TB2 %>% 
  mutate(Pop_Indígena = case_when(
    Raça %in% c("Indígena")~ "Indígena",
    TRUE ~ "Não"))


# Variável População vulnerável
# Criar várias colunas  combinando as populações e depois linkar em uma única coluna:
#Fonte: Boletim Epidemiológico, MS, 2025 página 32: Priorização para a categorização de acordo com o risco de adoecimento:
# Situação rua > Privado liberdade > HIV > Profissional saúde > Imigrante > Indígena


TB2 <- TB2 %>%
  mutate(
    OBSERVACAO1 = case_when(
      Situacao_rua == "Sim" ~ "Pessoa em situacao rua",
      TRUE ~ NA_character_
    )
  )
TB2 <- TB2 %>%
  mutate(
    OBSERVACAO2 = case_when(
      Privado_liberdade == "Sim" &
        (Situacao_rua %in% c("Não","Sem informação")) ~ "Pessoa Privada de liberdade",
      TRUE ~ NA_character_
    )
  )
TB2 <- TB2 %>%
  mutate(
    OBSERVACAO3 = case_when(
      PVHA == "Coinfecção TB/HIV" &
        (Situacao_rua %in% c("Não", "Sem informação")) &
        (Privado_liberdade %in% c("Não", "Sem informação")) ~ "Pessoa vivendo com HIV/aids",
      TRUE ~ NA_character_
    )
  )
TB2 <- TB2 %>%
  mutate(
    OBSERVACAO4 = case_when(
      Profissional_saude == "Sim" &
        (Situacao_rua %in% c("Não", "Sem informação")) &
        (Privado_liberdade %in% c("Não", "Sem informação")) &
        (PVHA %in% c("Ignorado","Infecção TB")) ~ "Profissional de Saúde",
      TRUE ~ NA_character_
    )
  )
TB2 <- TB2 %>%
  mutate(
    OBSERVACAO5 = case_when(
      Imigrante == "Sim" &
        (Situacao_rua %in% c("Não", "Sem informação")) &
        (Privado_liberdade %in% c("Não", "Sem informação")) &
        (PVHA %in% c("Ignorado","Não")) &
        (Profissional_saude %in% c("Não","Sem informação"))~ "População imigrante",
      TRUE ~ NA_character_
    )
  )
TB2 <- TB2 %>%
  mutate(
    OBSERVACAO6 = case_when(
      Pop_Indígena == "Indígena" &
        (Situacao_rua %in% c("Não", "Sem informação")) &
        (Privado_liberdade %in% c("Não", "Sem informação")) &
        (PVHA %in% c("Ignorado","Infecçao TB")) &
        (Profissional_saude %in% c("Não","Sem informação")) &
        (Imigrante %in% c("Não","Sem informação")) ~ "População indígena",
      TRUE ~ NA_character_
    )
  )

TB2 <- TB2 %>%
  mutate(
    OBSERVACAO7 = case_when(
        (Situacao_rua %in% c("Não", "Sem informação")) &
        (Privado_liberdade %in% c("Não", "Sem informação")) &
        (PVHA %in% c("Ignorado","Infecção TB")) &
        (Profissional_saude %in% c("Não","Sem informação")) &
        (Imigrante %in% c("Não","Sem informação")) &
        (Pop_Indígena %in% c("Não")) ~ "População não vulnerável",
      TRUE ~ NA_character_
    )
  )


#JUNTAR TODAS AS INFORMAÇÕES e excluir colunas desnecessárias:
TB2 <- TB2 %>%
  mutate(Pop_vulneraveis = coalesce(OBSERVACAO1,
                                    OBSERVACAO2,
                                    OBSERVACAO3,
                                    OBSERVACAO4,
                                    OBSERVACAO5,
                                    OBSERVACAO6,
                                    OBSERVACAO7))
#Acrescentar "sem informação" nas células NAs;
TB2 <- TB2 %>%
  mutate(Pop_vulneraveis = if_else(is.na(Pop_vulneraveis) | Pop_vulneraveis == "", 
                                   "Sem informação", 
                                   Pop_vulneraveis))

TB2_final<-TB2 %>% select(-OBSERVACAO1,-OBSERVACAO2,-OBSERVACAO3,-OBSERVACAO4,-OBSERVACAO5,-OBSERVACAO6,-OBSERVACAO7,
                          -EXTRAPU1_N,-EXTRAPU2_N,-EXTRAPUL_O,-AGRAVALCOO,-AGRAVDIABE,-AGRAVDOENC,-AGRAVOUTDE,
                          -AGRAVDROGA,-AGRAVTABAC)


## EXportação o para o painel TB
#write.csv(TB2_final,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Demografia_vulnerabilidade.csv") 
#write.csv(TB2_final,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Demografia_vulnerabilidade.csv")




###################################################################  POPULAÇÕES VULNERÁVEIS  ########################################################

TB_vulneraveis<-TB2 %>% select(-OBSERVACAO1,-OBSERVACAO2,-OBSERVACAO3,-OBSERVACAO4,-OBSERVACAO5,-OBSERVACAO6,-OBSERVACAO7,-TRATAMENTO,
                    -SITUA_ENCE,-IDADEordem,-Gestante,,-Pop_vulneraveis,-AGRAVALCOO,-AGRAVDIABE,-AGRAVDOENC,-AGRAVOUTDE,-AGRAVDROGA,-AGRAVTABAC,
                    -EXTRAPU1_N,-EXTRAPU2_N,-EXTRAPUL_O, -FORMA)

###   Juntar as variáveis "PVHA","Privado_liberdade", "Situacao_rua","Imigrante","Profissional_saude","Pop_Indígena" 
#na variável Vulnerabilidade

## Trocar o nome das categorias
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop1 = case_when(
    PVHA %in% c("Coinfecção TB/HIV")~ "PVHA",
    PVHA %in% c("Infecção TB")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-PVHA)
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop2 = case_when(
    Privado_liberdade %in% c("Sim")~ "Pop. privada de liberdade",
    Privado_liberdade %in% c("Não")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-Privado_liberdade)
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop3 = case_when(
    Situacao_rua %in% c("Sim")~ "Pop. em situação de rua",
    Situacao_rua %in% c("Não")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-Situacao_rua)
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop4 = case_when(
    Imigrante %in% c("Sim")~ "Pop. imigrante",
    Imigrante %in% c("Não")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-Imigrante)
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop5 = case_when(
    Profissional_saude %in% c("Sim")~ "Profissional de saúde",
    Profissional_saude %in% c("Não")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-Profissional_saude)
TB_vulneraveis<-TB_vulneraveis %>% 
  mutate(TB_Pop6 = case_when(
    Pop_Indígena %in% c("Indígena")~ "População Indígena",
    Pop_Indígena %in% c("Não")~ "Não",
    TRUE ~ "Sem informação"))
TB_vulneraveis<-TB_vulneraveis %>% select(-Pop_Indígena)


# Juntar todas  variáveis na variável "Vulnerabilidade" :
TB_vulneraveis <- TB_vulneraveis %>%
  pivot_longer(cols = TB_Pop1:TB_Pop6,  values_to = "Vulnerabilidade", values_drop_na = TRUE) %>% #exclui as linhas sem informação
  arrange(NU_NOTIFIC)  # Ordena o resultado
#Excluir coluna "name"criada pelo comando pivot_longer e as linhas repetidas
TB_vulneraveis<-TB_vulneraveis %>% select(-name)
TB_vulneraveis<-unique(TB_vulneraveis)

#EXcluir as linhas com a categoria não que têm  outra linha da mesma notificação com um tipo de vulnerabilidade: 
TB_vulneraveis <- TB_vulneraveis %>%
  group_by(NU_NOTIFIC) %>%
  filter(!(Vulnerabilidade == "Não" & n_distinct(Vulnerabilidade) > 1)) %>%
  ungroup()



## EXportação o para o painel TB
#write.csv(TB_vulneraveis,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/TB_vulneraveis.csv") 
#write.csv(TB_vulneraveis,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/TB_vulneraveis.csv")



###########################################     FORMAS EXTRAPULMONARES E COMORBIDADES/AGRAVOS ASSOCIADOS   ########################################################

## Seleção das variáveis:
TB3<-TB2 %>% select(-OBSERVACAO1,-OBSERVACAO2,-OBSERVACAO3,-OBSERVACAO4,-OBSERVACAO5,-OBSERVACAO6,-OBSERVACAO7,-TRATAMENTO,
                    -SITUA_ENCE,-IDADEordem,-Escolaridade,-Gestante,-Privado_liberdade,-Situacao_rua,-Imigrante,-Profissional_saude,
                    -Pop_Indígena,-Pop_vulneraveis,-AGRAVALCOO,-AGRAVDIABE,-AGRAVDOENC,-AGRAVOUTDE,-AGRAVDROGA,-AGRAVTABAC)

#Excluir as formas pulmonares e dados não informados:
TB3<-subset(TB3,FORMA!="Pulmonar" & FORMA!="Dado não informado")
TB3<-TB3 %>% select(-FORMA)


###   Variável TB_extrapulmonar: somar todas as informações das variáveis EXTRAPU1_N,EXTRAPU2_N e EXTRAPU_0:

## Tratar variáveis EXTRAPU1_N,EXTRAPU2_N e EXTRAPU_0: 
#EXTRAPU1_N:
TB3<-TB3 %>% 
  mutate(TB_EXtra1 = case_when(
    EXTRAPU1_N %in% c("01","1")~ "Pleural",
    EXTRAPU1_N %in% c("02","2")~ "Ganglionar",
    EXTRAPU1_N %in% c("3")~ "Geniturinária",
    EXTRAPU1_N %in% c("4")~ "Óssea",
    EXTRAPU1_N %in% c("5")~ "Ocular",
    EXTRAPU1_N %in% c("6")~ "Miliar",
    EXTRAPU1_N %in% c("07","7")~ "Meningoencefálica",
    EXTRAPU1_N %in% c("8")~ "Cutânea",
    EXTRAPU1_N %in% c("9")~ "Laríngea",
    EXTRAPU1_N %in% c("10")~ "Outra",
    TRUE ~ "Sem informação"))
TB3<-TB3 %>% select(-EXTRAPU1_N)

#EXTRAPU2_N:
TB3<-TB3 %>% 
  mutate(TB_EXtra2 = case_when(
    EXTRAPU2_N %in% c("1")~ "Pleural",
    EXTRAPU2_N %in% c("2")~ "Ganglionar",
    EXTRAPU2_N %in% c("3")~ "Geniturinária",
    EXTRAPU2_N %in% c("4")~ "Óssea",
    EXTRAPU2_N %in% c("5")~ "Ocular",
    EXTRAPU2_N %in% c("6")~ "Miliar",
    EXTRAPU2_N %in% c("07","7")~ "Meningoencefálica",
    EXTRAPU2_N %in% c("8")~ "Cutânea",
    EXTRAPU2_N %in% c("9")~ "Laríngea",
    EXTRAPU2_N %in% c("10")~ "Outra",
    TRUE ~ "Sem informação"))
TB3<-TB3 %>% select(-EXTRAPU2_N)


#EXTRAPUL_O 
#ATUALIZAR O COMANDO PERIODICAMENTE, avaliar posteriormente a criaçaõ da categoria: Abdominal =  fígado+intestino, etc...
TB3 <- TB3 %>% 
  mutate(
    EXTRAPUL_O = case_when(
      
      EXTRAPUL_O %in% c("MUSCULAR PLEURAL","ENFIZEMA PLEURAL","PERICARDICO + PLEURAL","PLEURAL",
                        "PLEURAL E CUTANEA","PLEURAL E GANGLIONAR GASTROINT","PLEURAL E MEINGOENCEFALICO",
                        "PLEURAL E MILIAR","PLEURAL/GENITURINARIA","PLEURAL/OSSEA","PLEUROPULMONAR","INTESTINAL,PELURAL",
                        "PERI+PLE") ~ "Pleural",
      
      EXTRAPUL_O %in% c("GALNGLOONA","GANG TORAXICA","GANG UNEDIASTINAL","GANG. CENTRAL",
                        "GANGLIENAR","GANGLINAL CENTRAL","GANGLIO","GANGLIO CERVICAL","GANGLIO DODOMINAL",
                        "GANGLIO PULMONAR","GANGLIOMA","GANGLIONA ABDOMINAL","GANGLIONAR","GANGLIONAR E MILIAR",
                        "GANGLIONAR E PERITONEAL","GANGLIONAR MEDIASTINAL","GANGLIONAR TORACICA","GANGLIONAR TORACICO",
                        "GANGLIOPERIFERICO","GANGLIOR","GANGLIOS CENTRAL","GANGLIOS INTRATORACICOS","GANGLIOS MEDIASTINO",
                        "GANGLIOS MESENTERICOS","GANGLIOTORACICO","GANGRIONAR ABDOMINAL","GNGLIO MEDIASTINAL",
                        "HEPATICA + GANGLIONAR","LEIFONADAL","LINDONODAL","LINF.ABDO","LINFADENITE","LINFADENITE CRONICA GRANULOMAT",
                        "LINFADENOPATIA","LINFATICA","LINFATICO","LINFATICO","LINFLADENITE","LINFOADENOPATIA TUBERCULOSA ME",
                        "LINFODENIDE","LINFODONA","LINFODONAL","LINFONODAL +","LINFONODAL AXILAR D","LINFONODO",
                        "LINFONODO CERVICAL","LINFONODO MEDIASTINAL","LINFONODO PARATRAQUEAL","LINFONODOS","LINFONOFIO CERVICAL",
                        "LOINFONODOS","PERITONEAL E GANGLIOS MESENTER","PULMAO,RENAL,GANGLIONAR","TB DAS GANGLIOS INTRATORACICOS",
                        "TB DO INTESTINO PERITONIO E GA","TB INTESTINAL GANGLIONAR","TBC GANGLIONAR","N\xd3DULO NO PULM\xc3O",
                        "TGI,F\xcdGADO,LINFATICOS","ABDOMINAL,CUTANEA GANGLIONAR","TGI,F\xcdGADO,LINFATICOS","GLAGLIONAR MEDIASTINAL",
                      "GLANG. INTRATORAXICA","GLANGLIO CERVICAL","GLANGLIONAR INTESTINAL","GRANGLIONAR") ~ "Ganglionar",
      
      EXTRAPUL_O %in% c("BEXIGA","BOLSA ESCROTAL","EPIDIDIMO","EPIDIDISSO","ESCOTRAL","OVARIANA","OVARIO","PLEURAL/GENITURINARIA",
                        "PROSTATA","PROSTATICA","PULMAO,RENAL,GANGLIONAR","RENAL","RENAL HEPATICA","RENAL?","TBC DE ENDOMETRIO",
                        "TECTICULAR","TESTICULAR","TESTICULO","TESTICULO BILATERAL","TROMPAS","ULMAO,CEREBRO E GENITURIARIA",
                        "URINARIA","URINARIA GENITORINARIO","UTERINA","INSUF RENAL","LINFONODAL","LINFONADO MEDIAS FINAL") ~ "Geniturinária",
      
      EXTRAPUL_O %in% c("ABCESSO MM ILIACO E","COLUNA","COLUNA CERVIAL","COLUNA LOMBAR","COLUNA TORACICA",
                        "COLUNA TORAXICA","COLUNA VERTEBRAL","ARTICULAR","MAL DE POTT","MAL POTT","MASTOIDITE,TB DO OUVIDO",
                      "MEDIASTINAL","MEDIASTINO","NODAL MEDIASTINO","OSSEA","OSTEARTICULAR","OSTEOCARTICULAR",
                      "OSTEOMUSCULAR","PELVE","PELVICA","PLEURAL/OSSEA","PLEURAL+OSSEA","TB MEDIASTINAL","TBERCULOSE DE CERVICAL",
                      "VERTEBRAL","OSTEO ARTICULAR","OSTEOARTICULAR") ~ "Óssea",
      
      EXTRAPUL_O %in% c("OCULAR","PELE E OCULAR","UVEITE","UVEITE DIFUSA") ~ "Ocular",
      
      EXTRAPUL_O %in% c("DESSEMINADA","DESSIMINADA","DISSEMINADA","DISSEMINADA NEURAL","DISSEMINADDA",
                        "DISSEMINADO","DISSEMINADO?","DISSEMINADOR","DISSEMIRADSA","DISSERMINADA TAI E MEDULA",
                        "DISSIMIMADA","DISSIMINADA","DISEMINADA","GANGLIONAR E MILIAR","MILIAR","MILIAR E MENINGO",
                        "MILIAR E MENINGOENCEFALICA","MILIAR E MENINGOENCEFALICO","MILIAR MENINGO","MILIAR/INTESTINAL",
                        "NODULAR","NODULO PULMONAR","NODULOS PULMONARES","PLEURAL E MILIAR","SEPSE","SIDA,SEPSI PULMONAR",
                        "TB DISSEMINADA","TB/MAC DISSEMINADA","ULMAO,CEREBRO E GENITURIARIA","LINFATICO MILIAR",
                        "PERITONEAL/MILIAR","ASSEMINADA") ~ "Miliar",
      
      EXTRAPUL_O %in% c("CEREBRAL","ENCEFALICO","MENIGITE TUBERCULOSA","MENINGEA","MENINGITE TUBERCULOSA","MENINGO + SNC",
                        "MENINGOECEFALICA","MIELITE","MIELITE TRANSVERSA TUBERCULOSA","MILIAR E MENINGO","MILIAR E MENINGOENCEFALICA",
                        "MILIAR E MENINGOENCEFALICO","MILIAR MENINGO","NEURAL","NEURO","NEURO TB","NEURO TUBERCULOSE",
                        "NEUROLOGICA","NEUROTB","NEUROTUBERCULOSE","PLEURAL E MEINGOENCEFALICO","SNC","ULMAO,CEREBRO E GENITURIARIA") ~ "Meningoencefálica",
      
      EXTRAPUL_O %in% c("CULTANEA","CUTANEA","ABDOMINAL,CUTANEA GANGLIONAR","PELE") ~ "Cutânea",
      
      EXTRAPUL_O %in% c("AMIGDALA","AMIGDALIANE","LARINGEA") ~ "Laríngea",
      
      TRUE ~ EXTRAPUL_O
    )
  )

#Preencher as células vazias:
TB3 <- TB3 %>%
  mutate(
    EXTRAPUL_O = if_else(is.na(EXTRAPUL_O), "Sem informação", EXTRAPUL_O)
  )

#Agrupar os demais tipos de TB em "Outra"

TB3<-TB3 %>% 
  mutate(TB_EXtra3 = case_when(
    EXTRAPUL_O %in% c("Pleural")~ "Pleural",
    EXTRAPUL_O %in% c("Ganglionar")~ "Ganglionar",
    EXTRAPUL_O %in% c("Geniturinária")~ "Geniturinária",
    EXTRAPUL_O %in% c("Óssea")~ "Óssea",
    EXTRAPUL_O %in% c("Ocular")~ "Ocular",
    EXTRAPUL_O %in% c("Miliar")~ "Miliar",
    EXTRAPUL_O %in% c("Meningoencefálica")~ "Meningoencefálica",
    EXTRAPUL_O %in% c("Cutânea")~ "Cutânea",
    EXTRAPUL_O %in% c("Laríngea")~ "Laríngea",
    EXTRAPUL_O %in% c("Sem informação")~ "Sem informação",
    TRUE ~ "Outra"))
TB3<-TB3 %>% select(-EXTRAPUL_O)


# Juntar as 3 variáveis na variável "TB_extrapulmonar" :
TB3 <- TB3 %>%
  pivot_longer(cols = TB_EXtra1:TB_EXtra3,  values_to = "TB_extrapulmonar", values_drop_na = TRUE) %>% #exclui as linhas sem informação
  arrange(NU_NOTIFIC)  # Ordena o resultado
#Excluir coluna "name"criada pelo comando pivot_longer e as linhas repetidas
TB3<-TB3 %>% select(-name)
TB3<-unique(TB3)
#Excluir as células sem informação do tipo de TB_extrapulmonar:
TB3<-subset(TB3,TB_extrapulmonar!="Sem informação")


## EXportação o para o painel TB
#write.csv(TB3,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Tuberculose_extrapulmonar.csv") 
#write.csv(TB3,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Tuberculose_extrapulmonar.csv")





###########################################     COMORBIDADES/AGRAVOS ASSOCIADOS   ########################################################

## Seleção das variáveis:
TB4<-TB2 %>% select(-OBSERVACAO1,-OBSERVACAO2,-OBSERVACAO3,-OBSERVACAO4,-OBSERVACAO5,-OBSERVACAO6,-OBSERVACAO7,-TRATAMENTO,
                    -SITUA_ENCE,-IDADEordem,-Escolaridade,-Gestante,-Privado_liberdade,-Situacao_rua,-Imigrante,-Profissional_saude,
                    -Pop_Indígena,-Pop_vulneraveis,-EXTRAPU1_N,-EXTRAPU2_N,-EXTRAPUL_O, -FORMA)

## Variável Agravos_comorbidades
# Tratar as variáveis: 
#AGRAVOS
TB4<-TB4 %>% 
  mutate(Agravo1 = case_when(
    AGRAVALCOO %in% c("1")~ "Uso nocivo de álcool",
    AGRAVALCOO %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB4<-TB4 %>% select(-AGRAVALCOO)

TB4<-TB4 %>% 
  mutate(Agravo2 = case_when(
    AGRAVDROGA %in% c("1")~ "Uso de drogas",
    AGRAVDROGA %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB4<-TB4 %>% select(-AGRAVDROGA)

TB4<-TB4 %>% 
  mutate(Agravo3 = case_when(
    AGRAVTABAC %in% c("1")~ "Tabagismo",
    AGRAVTABAC %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB4<-TB4 %>% select(-AGRAVTABAC)


#DOENÇAS
TB4<-TB4 %>% 
  mutate(Agravo4 = case_when(
    AGRAVDIABE %in% c("1")~ "Diabetes",
    AGRAVDIABE %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB4<-TB4 %>% select(-AGRAVDIABE)

TB4<-TB4 %>% 
  mutate(Agravo5 = case_when(
    AGRAVDOENC %in% c("1")~ "Doença mental",
    AGRAVDOENC %in% c("2")~ "Não",
    TRUE ~ "Sem informação"))
TB4<-TB4 %>% select(-AGRAVDOENC)

TB4<-TB4 %>% 
  mutate(Agravo6 = case_when(
    PVHA %in% c("Coinfecção TB/HIV")~ "PVHA",
    PVHA %in% c("Infecção TB")~ "Não",
    TRUE ~ "Sem informação"))


#Variável AGRAVOUTDE campo livre; recatgorizar: ATUALIZAR PERIODICAMENTE
#Selecionar as categoria com mais de 10 ocorrências:
TB4filtrado <- TB4 %>%
  add_count(AGRAVOUTDE, name = "freq") %>%
  filter(freq > 10) %>%
  select(-freq)
TB4filtrado %>%
  count(AGRAVOUTDE, sort = TRUE)

TB4 <- TB4 %>% 
  mutate(
    AGRAVOUTDE = case_when(
      AGRAVOUTDE %in% c("TABAGISMO","TABAGISTA","EX TABAGISTA","DROGADICAO","FUMANTE") ~ "Tabagismo",
      AGRAVOUTDE %in% c("DROGAS","USUARIO DE DROGAS","DROGADITO","DROGADICAO","USUARIO DE CRACK","DEPENDENTE QUIMICO","CRACK","DROGAS ILICITAS",
                        "DEPENDENCIA QUIMICA","MACONHA","USUARIO CRACK","USO DE CRACK","USUARIO DROGAS") ~ "Uso de drogas",
      AGRAVOUTDE %in% c("DEPRESSAO","DEPRESS\xc3O","ALZHEIMER","DROGADI\xc7\xc3O","ESQUIZOFRENIA") ~ "Doença mental",
      AGRAVOUTDE %in% c("HIV") ~ "PVHA",
      TRUE ~ AGRAVOUTDE
    )
  )


#Preencher as células vazias:
TB4 <- TB4 %>%
  mutate(
    AGRAVOUTDE = if_else(is.na(AGRAVOUTDE), "Sem informação", AGRAVOUTDE)
  )

#Agrupar os demais tipos de TB em "Outra"

TB4<-TB4 %>% 
  mutate(Agravo7 = case_when(
    AGRAVOUTDE %in% c("Tabagismo")~ "Tabagismo",
    AGRAVOUTDE %in% c("Uso de drogas")~ "Uso de drogas",
    AGRAVOUTDE %in% c("Doença mental")~ "Doença mental",
    AGRAVOUTDE %in% c("PVHA")~ "PVHA",
    AGRAVOUTDE %in% c("Sem informação")~ "Sem informação",
    TRUE ~ "Outra"))
TB4<-TB4 %>% select(-AGRAVOUTDE)

# Juntar as 6 variáveis na variável "Agravos_comorbidades" :
TB4 <- TB4 %>%
  pivot_longer(cols = Agravo1:Agravo7,  values_to = "Agravos_comorbidades", values_drop_na = TRUE) %>% #exclui as linhas sem informação
  arrange(NU_NOTIFIC)  # Ordena o resultado
#Excluir coluna "name"criada pelo comando pivot_longer e as linhas repetidas
TB4<-TB4 %>% select(-name)
TB4<-unique(TB4)

#Excluir as linhas que tem a categoria "Não" e também tem uma comorbidade.
TB4 <- TB4 %>%
  group_by(NU_NOTIFIC) %>%
  filter(!(Agravos_comorbidades == "Não" & n_distinct(Agravos_comorbidades) > 1)) %>%
  ungroup()

## EXportação o para o painel TB
#write.csv(TB4,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Comorbidades_agravos.csv") 
#write.csv(TB4,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Comorbidades_agravos.csv")


###################################################  INDICADORES   #####################################################

##  Considerar caso novo, pós-óbito e não sabe e forma pulmonar, Pulmonar e extrapulmonarr##
#FONTE: CADERNO DE INDICADORES DA TUBERCULOSE TUBERCULOSE SENSÍVEL, TUBERCULOSE DROGARRESISTENTE E TRATAMENTO PREVENTIVO Brasília – DF 2024

TB_Indicadores1_2_4 <- subset(TB,
  TRATAMENTO %in% c("Caso novo", "Pós-óbito", "Não sabe") &
    FORMA %in% c("Pulmonar", "Pulmonar e extrapulmonar")
)



#######    INDICADORES 1 e 2    ######## EXCLUIR EXAMES RETIRADOS NA CONFIRMAÇÃO LABORATORIAL
TB_CONTATO_CURA<-TB_Indicadores1_2_4[,c("ID_MN_RESI","Ano","NU_CONTATO","NU_COMU_EX","DT_DIAG","FORMA","TRATAMENTO","SITUA_ENCE","BACILOSC_E",
                "BACILOSC_2","BACILOS_E2","CULTURA_ES", "CULTURA_OU","TEST_MOLEC")] 
 

#Confirmação laboratorial: CONSIDERAR APENAS "BACILOSC_E",BACILOSC_2, BACILOS_E2,CULTURA_ES,TEST_MOLEC , Vide Caderno Indicadores pagina 55
TB_CONTATO_CURA <- TB_CONTATO_CURA %>%
  mutate(
    Confirmacao = case_when(
      if_any(c(BACILOSC_E, BACILOSC_2, BACILOS_E2,
               CULTURA_ES,CULTURA_OU), ~ . == "1") |  TEST_MOLEC %in% c("1", "2") 
      ~ "Caso confirmado",
      TRUE ~ "Caso não confirmado"
    )
  )
TB_CONTATO_CURA<-subset(TB_CONTATO_CURA,Confirmacao=="Caso confirmado" ) 


MUNICIPIOS3<-MUNICIPIOS[,c("NM_MUNICIP","ID_MN_RESI","Unidade.Regional.de.Saúde")]
#Criação variável "Ano", criar um ano para cada município de 2010 a 2025
MUNICIPIOS3 <- MUNICIPIOS3 %>%
  mutate(key = 1) %>%
  crossing(Ano = 2010:2025) %>%
  select(-key)


###   INDICADOR 1:proporção de contatos examinados dos casos novos de TB pulmonar com confirmação laboratorial  ####
#Indicador: Proporção de contatos examinados de casos novos de tuberculose pulmonar com confirmação laboratorial.
#Meta contatos tb - Portaria GM/MS Nº 6.878, DE 17 DE abril DE 2025
#10.Meta: 70% dos contatos dos casos novos de tuberculose pulmonar com confirmação laboratorial examinados.

#SELECIONAR VARIÁVEIS:
TB_CONTATO<-TB_CONTATO_CURA[,c("ID_MN_RESI","NU_CONTATO","NU_COMU_EX","Ano")]

#Acrescentando, regionais e os municípios faltantes:
n_distinct((TB_CONTATO$ID_MN_RESI)) #827 municípios com notificações
TB_CONTATO<- merge(TB_CONTATO, MUNICIPIOS3, by = c("ID_MN_RESI","Ano"),all=TRUE)

#Colocar zero nas células sem informção:
TB_CONTATO <- TB_CONTATO %>%
  mutate(across(c(NU_CONTATO, NU_COMU_EX), ~ ifelse(is.na(.), 0, .)))

# Proporção de contatos examinados por município e por regional:
#Total de contatos identificados e contatos examinados por ano por município
TB_CONTATO <- TB_CONTATO %>%
  group_by(ID_MN_RESI,Ano) %>%
  mutate(CONTATOS_IDENTIFICADOS_ANO_Municip = sum(NU_CONTATO)) %>%
  ungroup()
TB_CONTATO <- TB_CONTATO %>%
  group_by(ID_MN_RESI,Ano) %>%
  mutate(CONTATOS_EXAMINADOS_ANO_Municip = sum(NU_COMU_EX)) %>%
  ungroup()
#Variável proporção de contatos examinados
TB_CONTATO <- TB_CONTATO %>%
  mutate(CONTATOS_EXAMINADOS_PROP_Municip = round((CONTATOS_EXAMINADOS_ANO_Municip / CONTATOS_IDENTIFICADOS_ANO_Municip) * 100, 1))
#Substituir NaN (not avaiable number: onde contatos examinados e contatos identificados foram zero) por zero
TB_CONTATO$CONTATOS_EXAMINADOS_PROP_Municip[is.nan(TB_CONTATO$CONTATOS_EXAMINADOS_PROP_Municip)] <- 0
#Substituir inf (infinito: onde  contatos identificados foram zero) por zero
TB_CONTATO$CONTATOS_EXAMINADOS_PROP_Municip[is.infinite(TB_CONTATO$CONTATOS_EXAMINADOS_PROP_Municip)] <- 0
#Ajustar para valor máximo de 100;
TB_CONTATO <- TB_CONTATO %>%
  mutate(CONTATOS_EXAMINADOS_PROP_Municip = pmin(CONTATOS_EXAMINADOS_PROP_Municip, 100))

# Proporção de contatos examinados por município e por regional:
#Total de contatos identificados e contatos examinados por ano por regional
TB_CONTATO <- TB_CONTATO %>%
  group_by(Unidade.Regional.de.Saúde,Ano) %>%
  mutate(CONTATOS_IDENTIFICADOS_ANO_URS = sum(NU_CONTATO)) %>%
  ungroup()
TB_CONTATO <- TB_CONTATO %>%
  group_by(Unidade.Regional.de.Saúde,Ano) %>%
  mutate(CONTATOS_EXAMINADOS_ANO_URS = sum(NU_COMU_EX)) %>%
  ungroup()
#Variável proporção de contatos examinados
TB_CONTATO <- TB_CONTATO %>%
  mutate(CONTATOS_EXAMINADOS_PROP_URS = round((CONTATOS_EXAMINADOS_ANO_URS / CONTATOS_IDENTIFICADOS_ANO_URS) * 100, 1))
#Substituir NaN (not avaiable number: onde contatos examinados e contatos identificados foram zero) por zero
TB_CONTATO$CONTATOS_EXAMINADOS_PROP_URS[is.nan(TB_CONTATO$CONTATOS_EXAMINADOS_PROP_URS)] <- 0
#Substituir inf (infinito: onde  contatos identificados foram zero) por zero
TB_CONTATO$CONTATOS_EXAMINADOS_PROP_URS[is.infinite(TB_CONTATO$CONTATOS_EXAMINADOS_PROP_URS)] <- 0
#Ajustar para valor máximo de 100;
TB_CONTATO <- TB_CONTATO %>%
  mutate(CONTATOS_EXAMINADOS_PROP_URS = pmin(CONTATOS_EXAMINADOS_PROP_URS, 100))

#Seleção das variáveis:
TB_CONTATO <- unique(TB_CONTATO[,c("Ano","ID_MN_RESI","NM_MUNICIP","Unidade.Regional.de.Saúde",
                                   "CONTATOS_EXAMINADOS_PROP_Municip","CONTATOS_EXAMINADOS_PROP_URS")]) 



###   INDICADOR 2:Proporção de cura de casos novos de tuberculose pulmonar com confirmação laboratorial   ###   
#Plano Nacional de Saúde (PNS) 2020-2023, 
#meta 77,5% de cura .

#SELECIONAR VARIÁVEIS:
TB_CURA<-TB_CONTATO_CURA[,c("ID_MN_RESI","SITUA_ENCE","Ano")]

# MANTER CASOS NÃO ENCERRADOS: EXCLUIR TBDR, MUDANÇA DE ESQUEMA E FALÊNCIA
TB_CURA<-subset(TB_CURA, SITUA_ENCE!="TB-DR" & SITUA_ENCE!="Mudança de Esquema" & SITUA_ENCE!="Falência" )

#TOTAL DE NOTIFICAÇÕES POR ANO POR MUNICÍPIO
TB_CURA <- TB_CURA %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(NOTIF_ano_Municip = n()) %>%
  ungroup()

#Notificações com encerramento por cura:
TB_CURA <- TB_CURA %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(Encerra_cura_Municip = sum(SITUA_ENCE=="Cura",na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com encerramento como cura:
TB_CURA <- TB_CURA %>%
  mutate(CURA_PROP_Municip = round((Encerra_cura_Municip / NOTIF_ano_Municip) * 100, 1))

#Acrescentando os municípios faltantes e atribuir indicador de cura = 0
TB_CURA<- merge(TB_CURA, MUNICIPIOS3, by = c("ID_MN_RESI","Ano"),all=TRUE) 
TB_CURA <- TB_CURA %>% select(-SITUA_ENCE)
TB_CURA<-unique(TB_CURA)

#Colocar zero nas células sem informção:
TB_CURA <- TB_CURA %>%
  mutate(across(c(NOTIF_ano_Municip, Encerra_cura_Municip,CURA_PROP_Municip), ~ ifelse(is.na(.), 0, .)))


#TOTAL DE NOTIFICAÇÕES POR ANO POR Unidade.Regional.de.Saúde
TB_CURA <- TB_CURA %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(NOTIF_ano_URS = sum(NOTIF_ano_Municip,na.rm = TRUE)) %>%
  ungroup()

#Notificações com encerramento por cura POR ANO POR Unidade.Regional.de.Saúde
TB_CURA <- TB_CURA %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(Encerra_cura_URS = sum(Encerra_cura_Municip,na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com encerramento como cura:
TB_CURA <- TB_CURA %>%
  mutate(CURA_PROP_URS = round((Encerra_cura_URS / NOTIF_ano_URS) * 100, 1))

#Seleção das variáveis:
TB_CURA <- TB_CURA[,c("Ano","ID_MN_RESI","NM_MUNICIP","Unidade.Regional.de.Saúde",
                                   "CURA_PROP_Municip","CURA_PROP_URS")] 



###  INDICADOR 3: Proporção de realização de cultura entre os casos novos de tuberculose pulmonar  ###
#Dividir o indicador 3 em dois:
#realização de cultura em casos novos
#realização de cultura em casos de retratamento: considerar na variável TRATAMENTO  as categorias "Recidiva", "Reingresso após abandono",


#SELECIONAR VARIÁVEIS: #NÃO CONSIDEREI A VARIÁVEL CULTURA_OU (somente um caso em 2019)
TB_CULTURA_caso_novo <- subset(TB,
                              TRATAMENTO %in% c("Caso novo", "Pós-óbito", "Não sabe") &
                                FORMA %in% c("Pulmonar", "Pulmonar e extrapulmonar")
)

TB_CULTURA_retratamento <- subset(TB,
                               TRATAMENTO %in% c("Recidiva", "Reingresso após abandono") &
                                 FORMA %in% c("Pulmonar", "Pulmonar e extrapulmonar")
)


TB_CULTURA_caso_novo<-TB_CULTURA_caso_novo[,c("ID_MN_RESI","CULTURA_ES","Ano")]
TB_CULTURA_retratamento<-TB_CULTURA_retratamento[,c("ID_MN_RESI","CULTURA_ES","Ano")]



#CULTURA EM CASOS NOVOS

#TOTAL DE NOTIFICAÇÕES POR ANO POR MUNICÍPIO
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(NOTIF_ano_Municip_caso_novo = n()) %>%
  ungroup()

#Notificações com cultura por ano, excluir categoria 4: não realizada e 3: em andamento:
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(Cultura_ano_Municip_caso_novo = sum(CULTURA_ES !="3"& CULTURA_ES !="4",na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com cultura realizada:
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  mutate(CULTURA_PROP_Municip_caso_novo = round((Cultura_ano_Municip_caso_novo / NOTIF_ano_Municip_caso_novo) * 100, 1))

#Seleção das variáveis:
TB_CULTURA_caso_novo <- unique(TB_CULTURA_caso_novo[,c("ID_MN_RESI","Ano","NOTIF_ano_Municip_caso_novo","Cultura_ano_Municip_caso_novo","CULTURA_PROP_Municip_caso_novo")])

#Acrescentando os municípios faltantes:
TB_CULTURA_caso_novo<- merge(TB_CULTURA_caso_novo, MUNICIPIOS3, by = c("ID_MN_RESI","Ano"),all=TRUE) 

#Colocar zero nas células sem informação:
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  mutate(across(c(NOTIF_ano_Municip_caso_novo, Cultura_ano_Municip_caso_novo ,CULTURA_PROP_Municip_caso_novo), ~ ifelse(is.na(.), 0, .)))

#TOTAL DE NOTIFICAÇÕES POR ANO POR Unidade.Regional.de.Saúde
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(NOTIF_ano_URS_caso_novo = sum(NOTIF_ano_Municip_caso_novo,na.rm = TRUE)) %>%
  ungroup()

#Notificações com encerramento por cura POR ANO POR Unidade.Regional.de.Saúde
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(Cultura_ano_URS_caso_novo = sum(Cultura_ano_Municip_caso_novo,na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com encerramento como cura:
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo %>%
  mutate(CULTURA_PROP_URS_caso_novo = round((Cultura_ano_URS_caso_novo / NOTIF_ano_URS_caso_novo) * 100, 1))

#Seleção das variáveis:
TB_CULTURA_caso_novo <- TB_CULTURA_caso_novo[,c("Ano","ID_MN_RESI","NM_MUNICIP","Unidade.Regional.de.Saúde",
                      "CULTURA_PROP_Municip_caso_novo","CULTURA_PROP_URS_caso_novo")] 



#PROPORÇÃO DE CULTURA EM RETRATAMENTOS

#TOTAL DE NOTIFICAÇÕES POR ANO POR MUNICÍPIO
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(NOTIF_ano_Municip_retratamento = n()) %>%
  ungroup()

#Notificações com cultura por ano, excluir categoria 4: não realizada e 3: em andamento
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(Cultura_ano_Municip_retratamento = sum(CULTURA_ES !="3" & CULTURA_ES !="4",na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com cultura realizada:
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  mutate(CULTURA_PROP_Municip_retratamento = round((Cultura_ano_Municip_retratamento / NOTIF_ano_Municip_retratamento) * 100, 1))

#Seleção das variáveis:
TB_CULTURA_retratamento <- unique(TB_CULTURA_retratamento[,c("ID_MN_RESI","Ano","NOTIF_ano_Municip_retratamento","Cultura_ano_Municip_retratamento","CULTURA_PROP_Municip_retratamento")])

#Acrescentando os municípios faltantes:
TB_CULTURA_retratamento<- merge(TB_CULTURA_retratamento, MUNICIPIOS3, by = c("ID_MN_RESI","Ano"),all=TRUE) 

#Colocar zero nas células sem informção:
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  mutate(across(c(NOTIF_ano_Municip_retratamento, Cultura_ano_Municip_retratamento ,CULTURA_PROP_Municip_retratamento), ~ ifelse(is.na(.), 0, .)))

#TOTAL DE NOTIFICAÇÕES POR ANO POR Unidade.Regional.de.Saúde
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(NOTIF_ano_URS_retratamento = sum(NOTIF_ano_Municip_retratamento,na.rm = TRUE)) %>%
  ungroup()

#Notificações com encerramento por cura POR ANO POR Unidade.Regional.de.Saúde
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(Cultura_ano_URS_retratamento = sum(Cultura_ano_Municip_retratamento,na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com encerramento como cura:
TB_CULTURA_retratamento <- TB_CULTURA_retratamento %>%
  mutate(CULTURA_PROP_URS_retratamento = round((Cultura_ano_URS_retratamento / NOTIF_ano_URS_retratamento) * 100, 1))

#Seleção das variáveis:
TB_CULTURA_retratamento <- TB_CULTURA_retratamento[,c("Ano","ID_MN_RESI","NM_MUNICIP","Unidade.Regional.de.Saúde",
                                                "CULTURA_PROP_Municip_retratamento","CULTURA_PROP_URS_retratamento")] 




###  INDICADOR 4: Proporção de casos novos de tuberculose pulmonar diagnosticados por teste rápido molecular (TRM-TB)


#SELECIONAR VARIÁVEIS:
TB_TRM<-TB_Indicadores1_2_4[,c("ID_MN_RESI","TEST_MOLEC","Ano")]

#TOTAL DE NOTIFICAÇÕES POR ANO POR MUNICÍPIO
TB_TRM<- TB_TRM %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(NOTIF_ano_Municip = n()) %>%
  ungroup()

#Notificacções com TRM por ano, não considerar categoria 5: não realizada e NAs:
TB_TRM <- TB_TRM %>%
  group_by(Ano,ID_MN_RESI) %>%
  mutate(TRM_ano_Municip = sum(TEST_MOLEC!="5",na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com cultura realizada:
TB_TRM <- TB_TRM %>%
  mutate(TRM_PROP_ano_Municip = round((TRM_ano_Municip / NOTIF_ano_Municip) * 100, 1))

#Seleção das variáveis:
TB_TRM<- unique(TB_TRM[,c("ID_MN_RESI","Ano","NOTIF_ano_Municip","TRM_ano_Municip","TRM_PROP_ano_Municip")])

#Acrescentando os municípios faltantes:
TB_TRM<- merge(TB_TRM, MUNICIPIOS3, by = c("ID_MN_RESI","Ano"),all=TRUE) 

#Colocar zero nas células sem informção:
TB_TRM <- TB_TRM %>%
  mutate(across(c(NOTIF_ano_Municip, TRM_ano_Municip,TRM_PROP_ano_Municip), ~ ifelse(is.na(.), 0, .)))

#TOTAL DE NOTIFICAÇÕES POR ANO POR Unidade.Regional.de.Saúde
TB_TRM <- TB_TRM %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(NOTIF_ano_URS = sum(NOTIF_ano_Municip,na.rm = TRUE)) %>%
  ungroup()

#Notificações com encerramento por cura POR ANO POR Unidade.Regional.de.Saúde
TB_TRM <- TB_TRM %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(TRM_ano_URS = sum(TRM_ano_Municip,na.rm = TRUE)) %>%
  ungroup()

#Proporção de notificações com encerramento como cura:
TB_TRM <- TB_TRM %>%
  mutate(TRM_PROP_ano_URS = round((TRM_ano_URS / NOTIF_ano_URS) * 100, 1))

#Seleção das variáveis:
TB_TRM <- TB_TRM[,c("Ano","ID_MN_RESI","NM_MUNICIP","Unidade.Regional.de.Saúde",
                            "TRM_PROP_ano_Municip","TRM_PROP_ano_URS")] 


######################   LINKAGEM DOS BANCOS   ########################
TB_INDICADORES_final<-merge(TB_CONTATO,TB_CURA, by=c("ID_MN_RESI","Ano","NM_MUNICIP","Unidade.Regional.de.Saúde"))
TB_INDICADORES_final<-merge(TB_INDICADORES_final,TB_CULTURA_caso_novo, by=c("ID_MN_RESI","Ano","NM_MUNICIP","Unidade.Regional.de.Saúde"))
TB_INDICADORES_final<-merge(TB_INDICADORES_final,TB_CULTURA_retratamento, by=c("ID_MN_RESI","Ano","NM_MUNICIP","Unidade.Regional.de.Saúde"))
TB_INDICADORES_final<-merge(TB_INDICADORES_final,TB_TRM, by=c("ID_MN_RESI","Ano","NM_MUNICIP","Unidade.Regional.de.Saúde"))



## EXportação o para o painel TB
#write.csv(TB_INDICADORES_final,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Indicadores.csv") 
#write.csv(TB_INDICADORES_final,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/Indicadores.csv")

## Exportação para planilha:
#write.xlsx(TB_INDICADORES_final,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/TB_INDICADORES.xlsx") 
#write.xlsx(TB_INDICADORES_final,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/TB_INDICADORES.xlsx") 



###  INDICADOR 5: Busca ativa de sintomáticos respiratórios NO PERÍODO DE 2016 A 2025


####    IMPORTAÇÃO E TRATAMENTOS INDIVIDUALIZADOS POR ANO####

## 2016

SR_2016_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Alfenas")
SR_2016_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Barbacena")
SR_2016_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "BH")
SR_2016_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Coronel_Fabriciano")
SR_2016_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Diamantina")
SR_2016_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Divinopolis")
SR_2016_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Gov_Valadares")
SR_2016_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Itabira")
SR_2016_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Ituiutaba")
SR_2016_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Januaria")
SR_2016_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "JF")
SR_2016_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Leopoldina")
SR_2016_Manhumirim<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Manhumirim")
SR_2016_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Montes_Claros")
SR_2016_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Passos")
SR_2016_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Patos_Minas")
SR_2016_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Pedra_Azul")
SR_2016_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Pirapora")
SR_2016_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Ponte_Nova")
SR_2016_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Pouso_Alegre")
SR_2016_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Sao_Joao_Del_Rei")
SR_2016_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Sete_Lagoas")
SR_2016_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Teofilo_Otoni")
SR_2016_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Uba")
SR_2016_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Uberaba")
SR_2016_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Uberlandia")
SR_2016_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Unai")
SR_2016_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2016.xls",sheet = "Varginha")

SR_2016_MG<-rbind(SR_2016_Alfenas,SR_2016_Barbacena,SR_2016_BH,SR_2016_Coronel_Fabriciano,SR_2016_Diamantina,SR_2016_Divinópolis,SR_2016_Gov_Valadares,
                  SR_2016_Itabira,SR_2016_Ituiutaba,SR_2016_Januaria,SR_2016_JF,SR_2016_Leopoldina,SR_2016_Manhumirim,SR_2016_Montes_Claros,SR_2016_Passos,
                  SR_2016_Patos_Minas,SR_2016_Pedra_Azul,SR_2016_Pirapora,SR_2016_Ponte_Nova,SR_2016_Pouso_Alegre,SR_2016_Sao_Joao_Del_Rei,SR_2016_Sete_Lagoas,
                  SR_2016_Teofilo_Otoni,SR_2016_Uba,SR_2016_Uberaba,SR_2016_Uberlandia,SR_2016_Unai,SR_2016_Varginha)

#Exclusão de colunas              
SR_2016_MG<-SR_2016_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2016_MG)[names(SR_2016_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2016_MG)[names(SR_2016_MG) == "...2"] <- "POP_Municip"
names(SR_2016_MG)[names(SR_2016_MG) == "...3"] <- "SR_Municip"
names(SR_2016_MG)[names(SR_2016_MG) == "...4"] <- "Janeiro"
names(SR_2016_MG)[names(SR_2016_MG) == "...5"] <- "Fevereiro"
names(SR_2016_MG)[names(SR_2016_MG) == "...6"] <- "Março"
names(SR_2016_MG)[names(SR_2016_MG) == "...7"] <- "Abril"
names(SR_2016_MG)[names(SR_2016_MG) == "...8"] <- "Maio"
names(SR_2016_MG)[names(SR_2016_MG) == "...9"] <- "Junho"
names(SR_2016_MG)[names(SR_2016_MG) == "...10"] <- "Julho"
names(SR_2016_MG)[names(SR_2016_MG) == "...11"] <- "Agosto"
names(SR_2016_MG)[names(SR_2016_MG) == "...12"] <- "Setembro"
names(SR_2016_MG)[names(SR_2016_MG) == "...13"] <- "Outubro"
names(SR_2016_MG)[names(SR_2016_MG) == "...14"] <- "Novembro"
names(SR_2016_MG)[names(SR_2016_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2016_MG$Ano<-"2016"



## 2017

SR_2017_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Alfenas")
SR_2017_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Barbacena")
SR_2017_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "BH")
SR_2017_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Coronel_Fabriciano")
SR_2017_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Diamantina")
SR_2017_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Divinopolis")
SR_2017_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Gov_Valadares")
SR_2017_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Itabira")
SR_2017_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Ituiutaba")
SR_2017_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Januaria")
SR_2017_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "JF")
SR_2017_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Leopoldina")
SR_2017_Manhumirim<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Manhumirim")
SR_2017_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Montes_Claros")
SR_2017_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Passos")
SR_2017_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Patos_Minas")
SR_2017_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Pedra_Azul")
SR_2017_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Pirapora")
SR_2017_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Ponte_Nova")
SR_2017_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Pouso_Alegre")
SR_2017_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Sao_Joao_Del_Rei")
SR_2017_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Sete_Lagoas")
SR_2017_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Teofilo_Otoni")
SR_2017_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Uba")
SR_2017_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Uberaba")
SR_2017_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Uberlandia")
SR_2017_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Unai")
SR_2017_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2017.xls",sheet = "Varginha")

SR_2017_MG<-rbind(SR_2017_Alfenas,SR_2017_Barbacena,SR_2017_BH,SR_2017_Coronel_Fabriciano,SR_2017_Diamantina,SR_2017_Divinópolis,SR_2017_Gov_Valadares,
                  SR_2017_Itabira,SR_2017_Ituiutaba,SR_2017_Januaria,SR_2017_JF,SR_2017_Leopoldina,SR_2017_Manhumirim,SR_2017_Montes_Claros,SR_2017_Passos,
                  SR_2017_Patos_Minas,SR_2017_Pedra_Azul,SR_2017_Pirapora,SR_2017_Ponte_Nova,SR_2017_Pouso_Alegre,SR_2017_Sao_Joao_Del_Rei,SR_2017_Sete_Lagoas,
                  SR_2017_Teofilo_Otoni,SR_2017_Uba,SR_2017_Uberaba,SR_2017_Uberlandia,SR_2017_Unai,SR_2017_Varginha)

#Exclusão de colunas              
SR_2017_MG<-SR_2017_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2017_MG)[names(SR_2017_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2017_MG)[names(SR_2017_MG) == "POPULAÇÃO ESTIMADA P/ TCU 2016"] <- "POP_Municip"
names(SR_2017_MG)[names(SR_2017_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2017_MG)[names(SR_2017_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2017"] <- "Janeiro"
names(SR_2017_MG)[names(SR_2017_MG) == "...5"] <- "Fevereiro"
names(SR_2017_MG)[names(SR_2017_MG) == "...6"] <- "Março"
names(SR_2017_MG)[names(SR_2017_MG) == "...7"] <- "Abril"
names(SR_2017_MG)[names(SR_2017_MG) == "...8"] <- "Maio"
names(SR_2017_MG)[names(SR_2017_MG) == "...9"] <- "Junho"
names(SR_2017_MG)[names(SR_2017_MG) == "...10"] <- "Julho"
names(SR_2017_MG)[names(SR_2017_MG) == "...11"] <- "Agosto"
names(SR_2017_MG)[names(SR_2017_MG) == "...12"] <- "Setembro"
names(SR_2017_MG)[names(SR_2017_MG) == "...13"] <- "Outubro"
names(SR_2017_MG)[names(SR_2017_MG) == "...14"] <- "Novembro"
names(SR_2017_MG)[names(SR_2017_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2017_MG$Ano<-"2017"



## 2018

SR_2018_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Alfenas")
SR_2018_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Barbacena")
SR_2018_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "BH")
SR_2018_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Coronel_Fabriciano")
SR_2018_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Diamantina")
SR_2018_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Divinopolis")
SR_2018_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Gov_Valadares")
SR_2018_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Itabira")
SR_2018_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Ituiutaba")
SR_2018_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Januaria")
SR_2018_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "JF")
SR_2018_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Leopoldina")
SR_2018_Manhumirim<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Manhumirim")
SR_2018_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Montes_Claros")
SR_2018_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Passos")
SR_2018_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Patos_Minas")
SR_2018_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Pedra_Azul")
SR_2018_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Pirapora")
SR_2018_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Ponte_Nova")
SR_2018_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Pouso_Alegre")
SR_2018_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Sao_Joao_Del_Rei")
SR_2018_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Sete_Lagoas")
SR_2018_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Teofilo_Otoni")
SR_2018_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Uba")
SR_2018_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Uberaba")
SR_2018_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Uberlandia")
SR_2018_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Unai")
SR_2018_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2018.xls",sheet = "Varginha")

SR_2018_MG<-rbind(SR_2018_Alfenas,SR_2018_Barbacena,SR_2018_BH,SR_2018_Coronel_Fabriciano,SR_2018_Diamantina,SR_2018_Divinópolis,SR_2018_Gov_Valadares,
                  SR_2018_Itabira,SR_2018_Ituiutaba,SR_2018_Januaria,SR_2018_JF,SR_2018_Leopoldina,SR_2018_Manhumirim,SR_2018_Montes_Claros,SR_2018_Passos,
                  SR_2018_Patos_Minas,SR_2018_Pedra_Azul,SR_2018_Pirapora,SR_2018_Ponte_Nova,SR_2018_Pouso_Alegre,SR_2018_Sao_Joao_Del_Rei,SR_2018_Sete_Lagoas,
                  SR_2018_Teofilo_Otoni,SR_2018_Uba,SR_2018_Uberaba,SR_2018_Uberlandia,SR_2018_Unai,SR_2018_Varginha)

#Exclusão de colunas              
SR_2018_MG<-SR_2018_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2018_MG)[names(SR_2018_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2018_MG)[names(SR_2018_MG) == "POPULAÇÃO ESTIMADA P/ TCU 2016"] <- "POP_Municip"
names(SR_2018_MG)[names(SR_2018_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2018_MG)[names(SR_2018_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2018"] <- "Janeiro"
names(SR_2018_MG)[names(SR_2018_MG) == "...5"] <- "Fevereiro"
names(SR_2018_MG)[names(SR_2018_MG) == "...6"] <- "Março"
names(SR_2018_MG)[names(SR_2018_MG) == "...7"] <- "Abril"
names(SR_2018_MG)[names(SR_2018_MG) == "...8"] <- "Maio"
names(SR_2018_MG)[names(SR_2018_MG) == "...9"] <- "Junho"
names(SR_2018_MG)[names(SR_2018_MG) == "...10"] <- "Julho"
names(SR_2018_MG)[names(SR_2018_MG) == "...11"] <- "Agosto"
names(SR_2018_MG)[names(SR_2018_MG) == "...12"] <- "Setembro"
names(SR_2018_MG)[names(SR_2018_MG) == "...13"] <- "Outubro"
names(SR_2018_MG)[names(SR_2018_MG) == "...14"] <- "Novembro"
names(SR_2018_MG)[names(SR_2018_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2018_MG$Ano<-"2018"



## 2019

SR_2019_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Alfenas")
SR_2019_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Barbacena")
SR_2019_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "BH")
SR_2019_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Coronel_Fabriciano")
SR_2019_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Diamantina")
SR_2019_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Divinopolis")
SR_2019_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Gov_Valadares")
SR_2019_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Itabira")
SR_2019_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Ituiutaba")
SR_2019_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Januaria")
SR_2019_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "JF")
SR_2019_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Leopoldina")
SR_2019_Manhumirim<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Manhumirim")
SR_2019_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Montes_Claros")
SR_2019_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Passos")
SR_2019_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Patos_Minas")
SR_2019_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Pedra_Azul")
SR_2019_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Pirapora")
SR_2019_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Ponte_Nova")
SR_2019_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Pouso_Alegre")
SR_2019_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Sao_Joao_Del_Rei")
SR_2019_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Sete_Lagoas")
SR_2019_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Teofilo_Otoni")
SR_2019_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Uba")
SR_2019_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Uberaba")
SR_2019_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Uberlandia")
SR_2019_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Unai")
SR_2019_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2019.xls",sheet = "Varginha")

SR_2019_MG<-rbind(SR_2019_Alfenas,SR_2019_Barbacena,SR_2019_BH,SR_2019_Coronel_Fabriciano,SR_2019_Diamantina,SR_2019_Divinópolis,SR_2019_Gov_Valadares,
                  SR_2019_Itabira,SR_2019_Ituiutaba,SR_2019_Januaria,SR_2019_JF,SR_2019_Leopoldina,SR_2019_Manhumirim,SR_2019_Montes_Claros,SR_2019_Passos,
                  SR_2019_Patos_Minas,SR_2019_Pedra_Azul,SR_2019_Pirapora,SR_2019_Ponte_Nova,SR_2019_Pouso_Alegre,SR_2019_Sao_Joao_Del_Rei,SR_2019_Sete_Lagoas,
                  SR_2019_Teofilo_Otoni,SR_2019_Uba,SR_2019_Uberaba,SR_2019_Uberlandia,SR_2019_Unai,SR_2019_Varginha)

#Exclusão de colunas              
SR_2019_MG<-SR_2019_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2019_MG)[names(SR_2019_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2019_MG)[names(SR_2019_MG) == "POPULAÇÃO ESTIMADA P/ TCU 2016"] <- "POP_Municip"
names(SR_2019_MG)[names(SR_2019_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2019_MG)[names(SR_2019_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2019"] <- "Janeiro"
names(SR_2019_MG)[names(SR_2019_MG) == "...5"] <- "Fevereiro"
names(SR_2019_MG)[names(SR_2019_MG) == "...6"] <- "Março"
names(SR_2019_MG)[names(SR_2019_MG) == "...7"] <- "Abril"
names(SR_2019_MG)[names(SR_2019_MG) == "...8"] <- "Maio"
names(SR_2019_MG)[names(SR_2019_MG) == "...9"] <- "Junho"
names(SR_2019_MG)[names(SR_2019_MG) == "...10"] <- "Julho"
names(SR_2019_MG)[names(SR_2019_MG) == "...11"] <- "Agosto"
names(SR_2019_MG)[names(SR_2019_MG) == "...12"] <- "Setembro"
names(SR_2019_MG)[names(SR_2019_MG) == "...13"] <- "Outubro"
names(SR_2019_MG)[names(SR_2019_MG) == "...14"] <- "Novembro"
names(SR_2019_MG)[names(SR_2019_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2019_MG$Ano<-"2019"



## 2020

SR_2020_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Alfenas")
SR_2020_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Barbacena")
SR_2020_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "BH")
SR_2020_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Coronel_Fabriciano")
SR_2020_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Diamantina")
SR_2020_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Divinopolis")
SR_2020_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Gov_Valadares")
SR_2020_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Itabira")
SR_2020_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Ituiutaba")
SR_2020_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Januaria")
SR_2020_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "JF")
SR_2020_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Leopoldina")
SR_2020_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Manhuaçu")
SR_2020_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Montes_Claros")
SR_2020_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Passos")
SR_2020_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Patos_Minas")
SR_2020_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Pedra_Azul")
SR_2020_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Pirapora")
SR_2020_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Ponte_Nova")
SR_2020_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Pouso_Alegre")
SR_2020_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Sao_Joao_Del_Rei")
SR_2020_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Sete_Lagoas")
SR_2020_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Teofilo_Otoni")
SR_2020_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Uba")
SR_2020_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Uberaba")
SR_2020_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Uberlandia")
SR_2020_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Unai")
SR_2020_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2020.xls",sheet = "Varginha")

SR_2020_MG<-rbind(SR_2020_Alfenas,SR_2020_Barbacena,SR_2020_BH,SR_2020_Coronel_Fabriciano,SR_2020_Diamantina,SR_2020_Divinópolis,SR_2020_Gov_Valadares,
                  SR_2020_Itabira,SR_2020_Ituiutaba,SR_2020_Januaria,SR_2020_JF,SR_2020_Leopoldina,SR_2020_Manhuaçu,SR_2020_Montes_Claros,SR_2020_Passos,
                  SR_2020_Patos_Minas,SR_2020_Pedra_Azul,SR_2020_Pirapora,SR_2020_Ponte_Nova,SR_2020_Pouso_Alegre,SR_2020_Sao_Joao_Del_Rei,SR_2020_Sete_Lagoas,
                  SR_2020_Teofilo_Otoni,SR_2020_Uba,SR_2020_Uberaba,SR_2020_Uberlandia,SR_2020_Unai,SR_2020_Varginha)

#Exclusão de colunas              
SR_2020_MG<-SR_2020_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2020_MG)[names(SR_2020_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2020_MG)[names(SR_2020_MG) == "POPULAÇÃO ESTIMADA P/ TCU 2018"] <- "POP_Municip"
names(SR_2020_MG)[names(SR_2020_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2020_MG)[names(SR_2020_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2020"] <- "Janeiro"
names(SR_2020_MG)[names(SR_2020_MG) == "...5"] <- "Fevereiro"
names(SR_2020_MG)[names(SR_2020_MG) == "...6"] <- "Março"
names(SR_2020_MG)[names(SR_2020_MG) == "...7"] <- "Abril"
names(SR_2020_MG)[names(SR_2020_MG) == "...8"] <- "Maio"
names(SR_2020_MG)[names(SR_2020_MG) == "...9"] <- "Junho"
names(SR_2020_MG)[names(SR_2020_MG) == "...10"] <- "Julho"
names(SR_2020_MG)[names(SR_2020_MG) == "...11"] <- "Agosto"
names(SR_2020_MG)[names(SR_2020_MG) == "...12"] <- "Setembro"
names(SR_2020_MG)[names(SR_2020_MG) == "...13"] <- "Outubro"
names(SR_2020_MG)[names(SR_2020_MG) == "...14"] <- "Novembro"
names(SR_2020_MG)[names(SR_2020_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2020_MG$Ano<-"2020"


## 2021

SR_2021_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Alfenas")
SR_2021_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Barbacena")
SR_2021_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "BH")
SR_2021_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Coronel_Fabriciano")
SR_2021_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Diamantina")
SR_2021_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Divinopolis")
SR_2021_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Gov_Valadares")
SR_2021_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Itabira")
SR_2021_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Ituiutaba")
SR_2021_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Januaria")
SR_2021_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "JF")
SR_2021_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Leopoldina")
SR_2021_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Manhuaçu")
SR_2021_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Montes_Claros")
SR_2021_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Passos")
SR_2021_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Patos_Minas")
SR_2021_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Pedra_Azul")
SR_2021_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Pirapora")
SR_2021_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Ponte_Nova")
SR_2021_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Pouso_Alegre")
SR_2021_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Sao_Joao_Del_Rei")
SR_2021_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Sete_Lagoas")
SR_2021_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Teofilo_Otoni")
SR_2021_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Uba")
SR_2021_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Uberaba")
SR_2021_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Uberlandia")
SR_2021_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Unai")
SR_2021_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2021.xls",sheet = "Varginha")

SR_2021_MG<-rbind(SR_2021_Alfenas,SR_2021_Barbacena,SR_2021_BH,SR_2021_Coronel_Fabriciano,SR_2021_Diamantina,SR_2021_Divinópolis,SR_2021_Gov_Valadares,
                  SR_2021_Itabira,SR_2021_Ituiutaba,SR_2021_Januaria,SR_2021_JF,SR_2021_Leopoldina,SR_2021_Manhuaçu,SR_2021_Montes_Claros,SR_2021_Passos,
                  SR_2021_Patos_Minas,SR_2021_Pedra_Azul,SR_2021_Pirapora,SR_2021_Ponte_Nova,SR_2021_Pouso_Alegre,SR_2021_Sao_Joao_Del_Rei,SR_2021_Sete_Lagoas,
                  SR_2021_Teofilo_Otoni,SR_2021_Uba,SR_2021_Uberaba,SR_2021_Uberlandia,SR_2021_Unai,SR_2021_Varginha)

#Exclusão de colunas              
SR_2021_MG<-SR_2021_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2021_MG)[names(SR_2021_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2021_MG)[names(SR_2021_MG) == "POPULAÇÃO ESTIMADA P/ TCU 2019"] <- "POP_Municip"
names(SR_2021_MG)[names(SR_2021_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2021_MG)[names(SR_2021_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2021"] <- "Janeiro"
names(SR_2021_MG)[names(SR_2021_MG) == "...5"] <- "Fevereiro"
names(SR_2021_MG)[names(SR_2021_MG) == "...6"] <- "Março"
names(SR_2021_MG)[names(SR_2021_MG) == "...7"] <- "Abril"
names(SR_2021_MG)[names(SR_2021_MG) == "...8"] <- "Maio"
names(SR_2021_MG)[names(SR_2021_MG) == "...9"] <- "Junho"
names(SR_2021_MG)[names(SR_2021_MG) == "...10"] <- "Julho"
names(SR_2021_MG)[names(SR_2021_MG) == "...11"] <- "Agosto"
names(SR_2021_MG)[names(SR_2021_MG) == "...12"] <- "Setembro"
names(SR_2021_MG)[names(SR_2021_MG) == "...13"] <- "Outubro"
names(SR_2021_MG)[names(SR_2021_MG) == "...14"] <- "Novembro"
names(SR_2021_MG)[names(SR_2021_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2021_MG$Ano<-"2021"


## 2022

SR_2022_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Alfenas")
SR_2022_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Barbacena")
SR_2022_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "BH")
SR_2022_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Coronel_Fabriciano")
SR_2022_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Diamantina")
SR_2022_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Divinopolis")
SR_2022_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Gov_Valadares")
SR_2022_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Itabira")
SR_2022_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Ituiutaba")
SR_2022_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Januaria")
SR_2022_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "JF")
SR_2022_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Leopoldina")
SR_2022_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Manhuaçu")
SR_2022_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Montes_Claros")
SR_2022_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Passos")
SR_2022_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Patos_Minas")
SR_2022_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Pedra_Azul")
SR_2022_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Pirapora")
SR_2022_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Ponte_Nova")
SR_2022_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Pouso_Alegre")
SR_2022_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Sao_Joao_Del_Rei")
SR_2022_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Sete_Lagoas")
SR_2022_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Teofilo_Otoni")
SR_2022_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Uba")
SR_2022_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Uberaba")
SR_2022_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Uberlandia")
SR_2022_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Unai")
SR_2022_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2022.xls",sheet = "Varginha")

SR_2022_MG<-rbind(SR_2022_Alfenas,SR_2022_Barbacena,SR_2022_BH,SR_2022_Coronel_Fabriciano,SR_2022_Diamantina,SR_2022_Divinópolis,SR_2022_Gov_Valadares,
                  SR_2022_Itabira,SR_2022_Ituiutaba,SR_2022_Januaria,SR_2022_JF,SR_2022_Leopoldina,SR_2022_Manhuaçu,SR_2022_Montes_Claros,SR_2022_Passos,
                  SR_2022_Patos_Minas,SR_2022_Pedra_Azul,SR_2022_Pirapora,SR_2022_Ponte_Nova,SR_2022_Pouso_Alegre,SR_2022_Sao_Joao_Del_Rei,SR_2022_Sete_Lagoas,
                  SR_2022_Teofilo_Otoni,SR_2022_Uba,SR_2022_Uberaba,SR_2022_Uberlandia,SR_2022_Unai,SR_2022_Varginha)

#Exclusão de colunas              
SR_2022_MG<-SR_2022_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2022_MG)[names(SR_2022_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2022_MG)[names(SR_2022_MG) == "POPULAÇÃO ESTIMADA 2020"] <- "POP_Municip"
names(SR_2022_MG)[names(SR_2022_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2022_MG)[names(SR_2022_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2022"] <- "Janeiro"
names(SR_2022_MG)[names(SR_2022_MG) == "...5"] <- "Fevereiro"
names(SR_2022_MG)[names(SR_2022_MG) == "...6"] <- "Março"
names(SR_2022_MG)[names(SR_2022_MG) == "...7"] <- "Abril"
names(SR_2022_MG)[names(SR_2022_MG) == "...8"] <- "Maio"
names(SR_2022_MG)[names(SR_2022_MG) == "...9"] <- "Junho"
names(SR_2022_MG)[names(SR_2022_MG) == "...10"] <- "Julho"
names(SR_2022_MG)[names(SR_2022_MG) == "...11"] <- "Agosto"
names(SR_2022_MG)[names(SR_2022_MG) == "...12"] <- "Setembro"
names(SR_2022_MG)[names(SR_2022_MG) == "...13"] <- "Outubro"
names(SR_2022_MG)[names(SR_2022_MG) == "...14"] <- "Novembro"
names(SR_2022_MG)[names(SR_2022_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2022_MG$Ano<-"2022"


## 2023

SR_2023_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Alfenas")
SR_2023_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Barbacena")
SR_2023_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "BH")
SR_2023_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Coronel_Fabriciano")
SR_2023_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Diamantina")
SR_2023_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Divinopolis")
SR_2023_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Gov_Valadares")
SR_2023_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Itabira")
SR_2023_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Ituiutaba")
SR_2023_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Januaria")
SR_2023_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "JF")
SR_2023_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Leopoldina")
SR_2023_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Manhuaçu")
SR_2023_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Montes_Claros")
SR_2023_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Passos")
SR_2023_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Patos_Minas")
SR_2023_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Pedra_Azul")
SR_2023_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Pirapora")
SR_2023_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Ponte_Nova")
SR_2023_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Pouso_Alegre")
SR_2023_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Sao_Joao_Del_Rei")
SR_2023_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Sete_Lagoas")
SR_2023_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Teofilo_Otoni")
SR_2023_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Uba")
SR_2023_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Uberaba")
SR_2023_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Uberlandia")
SR_2023_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Unai")
SR_2023_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2023.xls",sheet = "Varginha")

SR_2023_MG<-rbind(SR_2023_Alfenas,SR_2023_Barbacena,SR_2023_BH,SR_2023_Coronel_Fabriciano,SR_2023_Diamantina,SR_2023_Divinópolis,SR_2023_Gov_Valadares,
                  SR_2023_Itabira,SR_2023_Ituiutaba,SR_2023_Januaria,SR_2023_JF,SR_2023_Leopoldina,SR_2023_Manhuaçu,SR_2023_Montes_Claros,SR_2023_Passos,
                  SR_2023_Patos_Minas,SR_2023_Pedra_Azul,SR_2023_Pirapora,SR_2023_Ponte_Nova,SR_2023_Pouso_Alegre,SR_2023_Sao_Joao_Del_Rei,SR_2023_Sete_Lagoas,
                  SR_2023_Teofilo_Otoni,SR_2023_Uba,SR_2023_Uberaba,SR_2023_Uberlandia,SR_2023_Unai,SR_2023_Varginha)

#Exclusão de colunas              
SR_2023_MG<-SR_2023_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2023_MG)[names(SR_2023_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2023_MG)[names(SR_2023_MG) == "POPULAÇÃO ESTIMADA 2021"] <- "POP_Municip"
names(SR_2023_MG)[names(SR_2023_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2023_MG)[names(SR_2023_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2023"] <- "Janeiro"
names(SR_2023_MG)[names(SR_2023_MG) == "...5"] <- "Fevereiro"
names(SR_2023_MG)[names(SR_2023_MG) == "...6"] <- "Março"
names(SR_2023_MG)[names(SR_2023_MG) == "...7"] <- "Abril"
names(SR_2023_MG)[names(SR_2023_MG) == "...8"] <- "Maio"
names(SR_2023_MG)[names(SR_2023_MG) == "...9"] <- "Junho"
names(SR_2023_MG)[names(SR_2023_MG) == "...10"] <- "Julho"
names(SR_2023_MG)[names(SR_2023_MG) == "...11"] <- "Agosto"
names(SR_2023_MG)[names(SR_2023_MG) == "...12"] <- "Setembro"
names(SR_2023_MG)[names(SR_2023_MG) == "...13"] <- "Outubro"
names(SR_2023_MG)[names(SR_2023_MG) == "...14"] <- "Novembro"
names(SR_2023_MG)[names(SR_2023_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2023_MG$Ano<-"2023"


## 2024

SR_2024_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Alfenas")
SR_2024_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Barbacena")
SR_2024_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "BH")
SR_2024_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Coronel_Fabriciano")
SR_2024_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Diamantina")
SR_2024_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Divinopolis")
SR_2024_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Gov_Valadares")
SR_2024_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Itabira")
SR_2024_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Ituiutaba")
SR_2024_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Januaria")
SR_2024_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "JF")
SR_2024_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Leopoldina")
SR_2024_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Manhuaçu")
SR_2024_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Montes_Claros")
SR_2024_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Passos")
SR_2024_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Patos_Minas")
SR_2024_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Pedra_Azul")
SR_2024_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Pirapora")
SR_2024_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Ponte_Nova")
SR_2024_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Pouso_Alegre")
SR_2024_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Sao_Joao_Del_Rei")
SR_2024_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Sete_Lagoas")
SR_2024_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Teofilo_Otoni")
SR_2024_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Uba")
SR_2024_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Uberaba")
SR_2024_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Uberlandia")
SR_2024_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Unai")
SR_2024_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2024.xls",sheet = "Varginha")

SR_2024_MG<-rbind(SR_2024_Alfenas,SR_2024_Barbacena,SR_2024_BH,SR_2024_Coronel_Fabriciano,SR_2024_Diamantina,SR_2024_Divinópolis,SR_2024_Gov_Valadares,
                  SR_2024_Itabira,SR_2024_Ituiutaba,SR_2024_Januaria,SR_2024_JF,SR_2024_Leopoldina,SR_2024_Manhuaçu,SR_2024_Montes_Claros,SR_2024_Passos,
                  SR_2024_Patos_Minas,SR_2024_Pedra_Azul,SR_2024_Pirapora,SR_2024_Ponte_Nova,SR_2024_Pouso_Alegre,SR_2024_Sao_Joao_Del_Rei,SR_2024_Sete_Lagoas,
                  SR_2024_Teofilo_Otoni,SR_2024_Uba,SR_2024_Uberaba,SR_2024_Uberlandia,SR_2024_Unai,SR_2024_Varginha)

#Exclusão de colunas              
SR_2024_MG<-SR_2024_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2024_MG)[names(SR_2024_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2024_MG)[names(SR_2024_MG) == "POPULAÇÃO ESTIMADA 2022"] <- "POP_Municip"
names(SR_2024_MG)[names(SR_2024_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2024_MG)[names(SR_2024_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2024"] <- "Janeiro"
names(SR_2024_MG)[names(SR_2024_MG) == "...5"] <- "Fevereiro"
names(SR_2024_MG)[names(SR_2024_MG) == "...6"] <- "Março"
names(SR_2024_MG)[names(SR_2024_MG) == "...7"] <- "Abril"
names(SR_2024_MG)[names(SR_2024_MG) == "...8"] <- "Maio"
names(SR_2024_MG)[names(SR_2024_MG) == "...9"] <- "Junho"
names(SR_2024_MG)[names(SR_2024_MG) == "...10"] <- "Julho"
names(SR_2024_MG)[names(SR_2024_MG) == "...11"] <- "Agosto"
names(SR_2024_MG)[names(SR_2024_MG) == "...12"] <- "Setembro"
names(SR_2024_MG)[names(SR_2024_MG) == "...13"] <- "Outubro"
names(SR_2024_MG)[names(SR_2024_MG) == "...14"] <- "Novembro"
names(SR_2024_MG)[names(SR_2024_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2024_MG$Ano<-"2024"


## 2025

SR_2025_Alfenas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Alfenas")
SR_2025_Barbacena<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Barbacena")
SR_2025_BH<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "BH")
SR_2025_Coronel_Fabriciano<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Coronel_Fabriciano")
SR_2025_Diamantina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Diamantina")
SR_2025_Divinópolis<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Divinopolis")
SR_2025_Gov_Valadares<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Gov_Valadares")
SR_2025_Itabira<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Itabira")
SR_2025_Ituiutaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Ituiutaba")
SR_2025_Januaria<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Januaria")
SR_2025_JF<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "JF")
SR_2025_Leopoldina<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Leopoldina")
SR_2025_Manhuaçu<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Manhuaçu")
SR_2025_Montes_Claros<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Montes_Claros")
SR_2025_Passos<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Passos")
SR_2025_Patos_Minas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Patos_Minas")
SR_2025_Pedra_Azul<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Pedra_Azul")
SR_2025_Pirapora<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Pirapora")
SR_2025_Ponte_Nova<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Ponte_Nova")
SR_2025_Pouso_Alegre<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Pouso_Alegre")
SR_2025_Sao_Joao_Del_Rei<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Sao_Joao_Del_Rei")
SR_2025_Sete_Lagoas<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Sete_Lagoas")
SR_2025_Teofilo_Otoni<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Teofilo_Otoni")
SR_2025_Uba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Uba")
SR_2025_Uberaba<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Uberaba")
SR_2025_Uberlandia<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Uberlandia")
SR_2025_Unai<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Unai")
SR_2025_Varginha<-read_excel("C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/Bancos_brutos/Sintomaticos_respiratorios/SR-Metas e Resultados_2025.xls",sheet = "Varginha")

SR_2025_MG<-rbind(SR_2025_Alfenas,SR_2025_Barbacena,SR_2025_BH,SR_2025_Coronel_Fabriciano,SR_2025_Diamantina,SR_2025_Divinópolis,SR_2025_Gov_Valadares,
                  SR_2025_Itabira,SR_2025_Ituiutaba,SR_2025_Januaria,SR_2025_JF,SR_2025_Leopoldina,SR_2025_Manhuaçu,SR_2025_Montes_Claros,SR_2025_Passos,
                  SR_2025_Patos_Minas,SR_2025_Pedra_Azul,SR_2025_Pirapora,SR_2025_Ponte_Nova,SR_2025_Pouso_Alegre,SR_2025_Sao_Joao_Del_Rei,SR_2025_Sete_Lagoas,
                  SR_2025_Teofilo_Otoni,SR_2025_Uba,SR_2025_Uberaba,SR_2025_Uberlandia,SR_2025_Unai,SR_2025_Varginha)

#Exclusão de colunas              
SR_2025_MG<-SR_2025_MG[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]

#Renomeação das colunas:
names(SR_2025_MG)[names(SR_2025_MG) == "MUNICÍPIO"] <- "NM_MUNICIP"
names(SR_2025_MG)[names(SR_2025_MG) == "POPULAÇÃO ESTIMADA 2024"] <- "POP_Municip"
names(SR_2025_MG)[names(SR_2025_MG) == "SR ESTIMADO/ANO: 1% da POP"] <- "SR_Municip"
names(SR_2025_MG)[names(SR_2025_MG) == "SINTOMÁTICOS RESPIRATÓRIOS (SR) EXAMINADOS NOS MUNICÍPIOS/MÊS EM 2025"] <- "Janeiro"
names(SR_2025_MG)[names(SR_2025_MG) == "...5"] <- "Fevereiro"
names(SR_2025_MG)[names(SR_2025_MG) == "...6"] <- "Março"
names(SR_2025_MG)[names(SR_2025_MG) == "...7"] <- "Abril"
names(SR_2025_MG)[names(SR_2025_MG) == "...8"] <- "Maio"
names(SR_2025_MG)[names(SR_2025_MG) == "...9"] <- "Junho"
names(SR_2025_MG)[names(SR_2025_MG) == "...10"] <- "Julho"
names(SR_2025_MG)[names(SR_2025_MG) == "...11"] <- "Agosto"
names(SR_2025_MG)[names(SR_2025_MG) == "...12"] <- "Setembro"
names(SR_2025_MG)[names(SR_2025_MG) == "...13"] <- "Outubro"
names(SR_2025_MG)[names(SR_2025_MG) == "...14"] <- "Novembro"
names(SR_2025_MG)[names(SR_2025_MG) == "...15"] <- "Dezembro"

#Ano:
SR_2025_MG$Ano<-"2025"



### LINKAR OS BANCOS PARA TRATAMENTO GERAL:
SR_16_25_MG<-rbind(SR_2016_MG,SR_2017_MG,SR_2018_MG,SR_2019_MG,SR_2020_MG,SR_2021_MG,SR_2022_MG,SR_2023_MG,SR_2024_MG,SR_2025_MG)

#Excluir cada linha da coluna "MUNICÍPIO" com o total da regional e com dado faltante
SR_16_25_MG <- subset(
  SR_16_25_MG,
  !is.na(NM_MUNICIP) & !NM_MUNICIP %in% c("SRS", "URS")
)

#Excluir caracteres especiais e substituir valores faltantes por zero:
SR_16_25_MG <- SR_16_25_MG %>%
  mutate(across(where(is.character), ~ gsub("[\\?\\*]", "", .x)))
SR_16_25_MG[is.na(SR_16_25_MG)] <- "0"


#Formatação de colunas 
SR_16_25_MG$SR_Municip<-as.integer(SR_16_25_MG$SR_Municip)
SR_16_25_MG$POP_Municip<-as.integer(SR_16_25_MG$POP_Municip)
SR_16_25_MG$Janeiro<-as.integer(SR_16_25_MG$Janeiro)
SR_16_25_MG$Fevereiro<-as.integer(SR_16_25_MG$Fevereiro)
SR_16_25_MG$Março<-as.integer(SR_16_25_MG$Março)
SR_16_25_MG$Abril<-as.integer(SR_16_25_MG$Abril)
SR_16_25_MG$Maio<-as.integer(SR_16_25_MG$Maio)
SR_16_25_MG$Junho<-as.integer(SR_16_25_MG$Junho)
SR_16_25_MG$Julho<-as.integer(SR_16_25_MG$Julho)
SR_16_25_MG$Agosto<-as.integer(SR_16_25_MG$Agosto)
SR_16_25_MG$Setembro<-as.integer(SR_16_25_MG$Setembro)
SR_16_25_MG$Outubro<-as.integer(SR_16_25_MG$Outubro)
SR_16_25_MG$Novembro<-as.integer(SR_16_25_MG$Novembro)
SR_16_25_MG$Dezembro<-as.integer(SR_16_25_MG$Dezembro)

#Corrigir nomes de municípios:
SR_16_25_MG$NM_MUNICIP <- gsub("Conc. Ipanema","Conceição de Ipanema",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("Dona Eusébia","Dona Euzébia",SR_16_25_MG$NM_MUNICIP) 
SR_16_25_MG$NM_MUNICIP <- gsub("Gouveia","Gouvêa",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("Olhos-d'Água" ,"Olhos-D'água",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("Pingo-d'Água" ,"Pingo-D'água",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("Queluzito","Queluzita",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("S.J. Mantimento","São José do Mantimento",SR_16_25_MG$NM_MUNICIP) 
SR_16_25_MG$NM_MUNICIP <- gsub("Santana  Mçu","Santana do Manhuaçu",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("São João del Rei","São João Del Rei",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("São João Mçu","São João do Manhuaçu",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("Santa Rita do Ibitipoca","Santa Rita de Ibitipoca",SR_16_25_MG$NM_MUNICIP)
SR_16_25_MG$NM_MUNICIP <- gsub("São Tomé das Letras","São Thomé das Letras",SR_16_25_MG$NM_MUNICIP)


#INCLUSÃO DE VARIÁVEIS:

#URS: de acordo com o PDR 2024:
MUNICIPIOS_SR<-MUNICIPIOS[,c("NM_MUNICIP","Unidade.Regional.de.Saúde")]
SR_16_25_MG<- merge(SR_16_25_MG, MUNICIPIOS_SR, by = c("NM_MUNICIP"),all=TRUE)

#População de SR por Unidade.Regional.de.Saúde 
SR_16_25_MG <- SR_16_25_MG %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(SR_URS = sum(SR_Municip,na.rm = TRUE)) %>%
  ungroup()

#Total de sintomáticos respiratórios examinados, por Município e por Unidade.Regional.de.Saúde :
SR_16_25_MG <- SR_16_25_MG %>%
  group_by(Ano) %>%
  mutate(
    SR_MUNICIP_Exam = rowSums(across(Janeiro:Dezembro), na.rm = TRUE))%>%
  ungroup()
SR_16_25_MG <- SR_16_25_MG %>%
  group_by(Ano,Unidade.Regional.de.Saúde) %>%
  mutate(SR_URS_Exam = sum(SR_MUNICIP_Exam,na.rm = TRUE)) %>%
  ungroup()

#Proporção de sintomáticos respiratórios examinados por município e por URS:
SR_16_25_MG <- SR_16_25_MG %>%
  mutate(PROP_SR_MUNICIP = round((SR_MUNICIP_Exam/SR_Municip) * 100, 0))
SR_16_25_MG <- SR_16_25_MG %>%
  mutate(PROP_SR_URS = round((SR_URS_Exam/SR_URS) * 100, 0))
#considerar até 100%
SR_16_25_MG$PROP_SR_MUNICIP[SR_16_25_MG$PROP_SR_MUNICIP > 100] <- 100
SR_16_25_MG$PROP_SR_URS[SR_16_25_MG$PROP_SR_URS > 100] <- 100



# Metas:
SR_16_25_MG$Meta <- ifelse(SR_16_25_MG$Ano == "2016", "40%",
                           ifelse(SR_16_25_MG$Ano <= "2020", "50%",
                                  ifelse(SR_16_25_MG$Ano == "2021", "60%",
                                         ifelse(SR_16_25_MG$Ano <= "2025", "65%",
                                                "65%"))))


## SELEÇÃO DAS VARIÁVEIS PARA O PAINEL: 
SR_FINAL_PAINEL<-SR_16_25_MG[,c("NM_MUNICIP","Ano","Unidade.Regional.de.Saúde","PROP_SR_MUNICIP",
                         "PROP_SR_URS","Meta")]




## EXportação o para o painel TB
#write.csv(SR_FINAL_PAINEL,"C:/Users/Cássia/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/SR_INDICADORES_final.csv") 
#write.csv(SR_FINAL_PAINEL,"C:/Users/x3996631/OneDrive - Secretaria de Estado de Saude de Minas Gerais/TBC/Painel/SR_INDICADORES_final.csv")




