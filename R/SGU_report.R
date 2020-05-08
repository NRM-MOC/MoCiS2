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
moc_write_SGU <- function(data, sheet, file, mall_path =  system.file("extdata", "miljogifter-leveransmall.xlsx", package = "MoCiS2"), program = "none"){
  options(scipen = 999)
  mall <- readxl::read_excel(mall_path, sheet = sheet)
  if (program %in% c("hav", "limn"))
    data <- data %>%
    mutate(PROVTAG_SYFTE = "NMO",
           PROVPLATS_TYP = "Bakgrund",
           PROVTAG_ORG = "NRM",
           ACKR_PROV = "Nej",
           DIREKT_BEHA = "FRYST",
           PROVPLATS_MILJO = ifelse(program == "hav", "HAV-BRACKV", "SJO-SOTV-RINN"),
           PLATTFORM = ifelse(program == "hav", "FISKEBAT", "SMABAT"),
           PLATTFORM = ifelse(ART == "Blamussla", "SAKNAS", PLATTFORM),
           PROVTAG_MET = ifelse(ART == "Blamussla", "Dykning", "Natfiske"),
           PROVTAG_MET = ifelse((ORGAN == "AGG") & (!is.na(ORGAN)), "Aggplockning", PROVTAG_MET),
           MATOSAKERHET_TYP = ifelse(is.na(MATOSAKERHET), NA, "U2")
    )
  data <- data %>%
    arrange(PARAMETERNAMN) %>%
    mutate_if(is.numeric, ~as.character(.x) %>% str_replace("[.]", ",")) %>%
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
              ANTAL_DAGAR = max(ANTAL_DAGAR),
              KON = pool_sex(KON),
              ANTAL = unique(ANTAL),
              ALDR = mean_or_na(ALDR),
              TOTV = mean_or_na(TOTV),
              TOTL = mean_or_na(TOTL)) %>%
    rename(PROV_KOD_ORIGINAL = PROV_KOD_ORIGINAL_POOL)

  if (add_bio_pars){
    kodlista <- tribble(
      ~NRM_PARAMETERKOD, ~PARAMETERNAMN, ~UNIK_PARAMETERKOD, ~LABB, ~ENHET,
      "ALDR", "Ålder", "CH12/239", "NRM", "ar",
      "ALDRH", "Ålder (medelvärde)", "CH12/241", "NRM", "ar",
      "TOTL", "Längd", "CH12/161", "NRM", "cm",
      "TOTLH", "Längd (medelvärde)", "CH12/163", "NRM", "cm",
      "TOTV", "Vikt", "CH12/232", "NRM", "g",
      "TOTVH", "Vikt (medelvärde)", "CH12/234", "NRM", "g",
    )
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
      ungroup()
  }
  else
  {
    data <- analysdata
  }
  data %>% left_join(bio_pool_data, by = "PROV_KOD_ORIGINAL")
}
