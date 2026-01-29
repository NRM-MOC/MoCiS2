#' Title
#'
#' @param data
#' @param sheet
#' @param file
#' @param mall_path
#' @param program
#'
#' @return
#' @export
#'
#' @examples
moc_write_SGU <- function(data, sheet, file, mall_path =  system.file("extdata", "mall-inrapportering-datavardskap-miljogifter-2026.xlsx", package = "MoCiS2"), program = "none"){
  options(scipen = 999)
  mall <- readxl::read_excel(mall_path, sheet = sheet)
  if (program %in% c("hav", "limn"))
    data <- data %>%
    mutate(PROVTAG_SYFTE = "NMO",
           PROVPLATS_TYP = "Bakgrund",
           PROVTAG_ORG = "NRM",
           ACKR_PROV = "Nej",
           DIREKT_BEHA = "FRYST",
           PROVDATA_TYP = "BIOTA",
           PROVPLATS_MILJO = ifelse(program == "hav", "HAV-BRACKV", "SJO-SOTV-RINN"),
           PLATTFORM = ifelse(program == "hav", "FISKEBAT", "SMABAT"),
           PLATTFORM = ifelse(ART == "Blamussla", "SAKNAS", PLATTFORM),
           PLATTFORM = ifelse(ART %in% c("Fisktarna", "Sillgrissla", "Strandskata"),"SAKNAS",PLATTFORM),
           #PROVTAG_MET = ifelse(ART == "Blamussla", "Dykning", "Natfiske"),
           PROVTAG_MET = ifelse((ART == "Blamussla" & NAMN_PROVPLATS =='Kvädöfjärden'),"Bottenskrapa", ifelse((ART == "Blamussla" & NAMN_PROVPLATS =='Nidingen'),"Dykning", ifelse((ART == "Blamussla" & NAMN_PROVPLATS =='Fjällbacka'),"Havskrap", "Natfiske"))),
           PROVTAG_MET = ifelse(ART %in% c("Fisktarna", "Sillgrissla", "Strandskata"), "Aggplockning", PROVTAG_MET),
           MATOSAKERHET_TYP = ifelse(is.na(MATOSAKERHET), NA, "U2")
    )
  data <- data %>%
    arrange(PARAMETERNAMN) %>%
    mutate_if(is.numeric, ~round(.x, 5) %>% as.character() %>% str_replace("[.]", ",")) %>%
    mutate_all(as.character) %>%
    select(intersect(names(mall), names(data))) %>%
    distinct()
  not_available <- setdiff(names(mall), names(data))
  if (length(not_available) > 0)
    message(paste0("Unavailable columns: ", paste(not_available[-1], collapse = ", ")))
  openxlsx::write.xlsx(bind_rows(mall, data), file = file)
}

#' Title
#'
#' @param biodata
#' @param analysdata
#' @param add_bio_pars
#'
#' @return
#' @export
#'
#' @examples
moc_join_SGU <- function(biodata, analysdata, add_bio_pars = TRUE){
  pool_sex <- function(sex){
    if (all(is.na(sex)))
      return(NA)
    sexes <- na.omit(sex) %>% unique()
    ifelse(length(sexes) == 1, sexes, "X")
  }
  bio_pool_data <- map_df(unique(analysdata$PROV_KOD_ORIGINAL), unpool) %>%
    left_join(biodata, by = "PROV_KOD_ORIGINAL") %>%
    mutate(KON = if_else(str_detect(PROV_KOD_ORIGINAL_POOL, "-") & (KON %in% c("F", "M")), "X", KON)) %>%
    group_by(PROV_KOD_ORIGINAL_POOL) %>%
    summarise(PROVTAG_DAT = min(PROVTAG_DAT),
              PROVTAG_TID = NA,
              PROVTAG_SLUTDAT = max(ANTAL_DAGAR),
              PROVTAG_SLUTTID = NA,
              KON = pool_sex(KON),
              ANTAL = unique(ANTAL),
              ALDR = mean_or_na(ALDR),
              TOTV = mean_or_na(TOTV),
              TOTL = mean_or_na(TOTL)) %>%
    rename(PROV_KOD_ORIGINAL = PROV_KOD_ORIGINAL_POOL)

  if (add_bio_pars){
    kodlista <- tribble( # This codelist is deprecated
      ~NRM_PARAMETERKOD, ~PARAMETERNAMN, ~UNIK_PARAMETERKOD, ~LABB, ~ENHET, ~PROV_BERED, ~PROVKARL, ~ORGAN, ~ANALYS_INSTR,
      "ALDR", "Ålder", "CH12/239", "NRM", "ar", "EJ_REL", "EJ_REL", "HELKROPP", "Stereomikroskop", 
      "ALDRH", "Ålder (medelvärde)", "CH12/241", "NRM", "ar", "EJ_REL", "EJ_REL", "HELKROPP", "Stereomikroskop",  
      "TOTL", "Längd", "CH12/161", "NRM", "cm", "EJ_REL", "EJ_REL", "HELKROPP", "LINJAL",
      "TOTLH", "Längd (medelvärde)", "CH12/163", "NRM", "cm", "EJ_REL", "EJ_REL", "HELKROPP", "LINJAL", 
      "TOTV", "Vikt", "CH12/232", "NRM", "g", "EJ_REL", "EJ_REL", "HELKROPP", "VAG",
      "TOTVH", "Vikt (medelvärde)", "CH12/234", "NRM", "g", "EJ_REL", "EJ_REL", "HELKROPP", "VAG" 
    )
    kodlista <- readxl::read_excel(system.file("extdata", "codelist.xlsx", package = "MoCiS2"), sheet = "PARAMETRAR")  %>%
      select(NRM_PARAMETERKOD, PARAMETERNAMN, UNIK_PARAMETERKOD, LABB, UTFOR_LABB, ENHET, PROV_BERED, PROVKARL, ANALYS_INSTR, ANALYS_MET, ACKREDITERAD_MET, PROV_LAGR) %>% 
      mutate(ORGAN = "HELKROPP")
    bio_measurements <- bio_pool_data %>%
      select(PROV_KOD_ORIGINAL, ALDR, TOTV, TOTL) %>%
      pivot_longer(c("ALDR", "TOTV", "TOTL"), names_to = "NRM_PARAMETERKOD", values_to = "MATVARDETAL") %>%
      mutate(MATVARDETAL = ifelse(is.nan(MATVARDETAL), NA, MATVARDETAL),
             NRM_PARAMETERKOD = ifelse(str_detect(PROV_KOD_ORIGINAL, "-"), paste0(NRM_PARAMETERKOD, "H"), NRM_PARAMETERKOD)) %>%
      filter(!is.na(MATVARDETAL)) %>%
      left_join(kodlista, by = "NRM_PARAMETERKOD") 
    data <- bind_rows(analysdata, bio_measurements) %>% 
      group_by(PROV_KOD_ORIGINAL) %>% 
      fill(PROVPLATS_ID, NAMN_PROVPLATS, ART, DYNTAXA_TAXON_ID, .direction = "downup") %>% 
      ungroup() %>% 
      mutate(ORGAN = ifelse(str_sub(NRM_PARAMETERKOD, 1, 4) == "ALDR",
                            case_when(ART == "Gadda" ~ "CLEITRUM",
                                      ART %in% c("Stromming", "Sill") ~ "FJALL",
                                      ART %in% c("Tanglake", "Torsk", "Abborre", "Roding") ~ "OTOLIT"),
                            ORGAN),
             ORGAN = ifelse((str_sub(NRM_PARAMETERKOD, 1, 4) == "TOTL") & (ART == "Blamussla"), "SKAL",
                            ORGAN)
      )
  }
  else
  {
    data <- analysdata
  }
  data %>% left_join(bio_pool_data, by = "PROV_KOD_ORIGINAL")
}
