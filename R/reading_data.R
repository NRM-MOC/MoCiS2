
#' Given PROV_KOD_ORIGINAL, splits into a table of individual PROV_KOD_ORIGINAL and corresponding pooled PROV_KOD_ORIGINAL
#'
#' @param PROV_KOD_ORIGINAL An NRM ACCNR
#'
#' @return A tibble
#'
#'
#' @examples
#' unpool("C2016/00937-00939")
unpool <- function(PROV_KOD_ORIGINAL){
  if (str_detect(PROV_KOD_ORIGINAL, "-")){ #homogenat
    head <- str_extract(PROV_KOD_ORIGINAL, ".*/")
    first <- str_extract(PROV_KOD_ORIGINAL, "(?<=/)([^-]*)") %>% as.integer()
    last <- str_extract(PROV_KOD_ORIGINAL, "(?<=-)(.*)") %>% as.integer()
    all <- first:last %>% as.character()
    PROV_KOD_ORIGINAL_IND <- paste0(head, str_pad(all, width = 5, pad = "0"))
    tibble(PROV_KOD_ORIGINAL_POOL = PROV_KOD_ORIGINAL, PROV_KOD_ORIGINAL = PROV_KOD_ORIGINAL_IND, ANTAL = last - first + 1)
  }
  else
  {
    tibble(PROV_KOD_ORIGINAL_POOL = PROV_KOD_ORIGINAL, PROV_KOD_ORIGINAL = PROV_KOD_ORIGINAL, ANTAL = 1)
  }
}

#' Reads the results sheet from a lab-protocol
#'
#' @param path
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file <- function(path, sheet = "results"){
  readxl::read_excel(path, sheet = sheet, skip = 1, na = c("-99.99", "N/A")) %>%
    rename(PROV_KOD_ORIGINAL = ...1, PROV_KOD_LABB = ...2, GENUS = ...3, PROVPLATS_ANALYSMALL  = ...4, ORGAN = ...5) %>%
    filter(!is.na(PROV_KOD_ORIGINAL)) %>%
    mutate(PROV_KOD_LABB = as.character(PROV_KOD_LABB)) %>%
    mutate_if(is.character, ~str_replace(.x, "<", "-")) %>% # Check if "<" rather than "-" is used for LOQ
    mutate_at(-(1:5), as.numeric) %>%
    select(-contains("..."))%>%
    pivot_longer(-(PROV_KOD_ORIGINAL:ORGAN), names_to = "NRM_PARAMETERKOD", values_to = "MATVARDETAL")
}

#' Title
#'
#' @param path
#' @param variable
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file2 <- function(path, variable, sheet){
  readxl::read_excel(path, sheet = sheet, skip = 1, na = c("-99.99", "N/A")) %>%
    rename(PROV_KOD_ORIGINAL = ...1, PROV_KOD_LABB = ...2, GENUS = ...3, PROVPLATS_ANALYSMALL = ...4) %>%
    filter(!is.na(PROV_KOD_ORIGINAL)) %>%
    mutate(PROV_KOD_LABB = as.character(PROV_KOD_LABB)) %>%
    mutate_if(is.character, ~str_replace(.x, "<", "-")) %>% # Check if "<" rather than "-" is used for LOQ
    mutate_at(-(1:4), as.numeric) %>%
    select(-PROV_KOD_LABB, -GENUS, -PROVPLATS_ANALYSMALL, -contains("...")) %>%
    pivot_longer(-PROV_KOD_ORIGINAL, names_to = "NRM_PARAMETERKOD", values_to = variable)
}

#' Title
#'
#' @param path
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file_date <- function(path, sheet = "date of analysis"){
  fix_date <- Vectorize(function(x){
    x <- as.numeric(x)
    if (is.na(x))
      return(NA)
    if (x < 100000)
      as.Date(x, origin = "1899-12-30") %>% as.character() # Excel-date
    else
      as.Date(as.character(x), "%Y%m%d") %>% as.character()
  })
  readxl::read_excel(path, sheet = sheet, na = c("-99.99", "N/A")) %>%
    rename(PROV_KOD_ORIGINAL = 1, PROV_KOD_LABB = 2, GENUS = 3, PROVPLATS_ANALYSMALL = 4) %>%
    filter(!is.na(PROV_KOD_ORIGINAL), !(PROV_KOD_ORIGINAL == 0)) %>%
    mutate(PROV_KOD_LABB = as.character(PROV_KOD_LABB)) %>%
    mutate_at(-(1:4), as.numeric) %>%
    select(-PROV_KOD_LABB, -GENUS, -PROVPLATS_ANALYSMALL, -contains("...")) %>%
    filter(PROV_KOD_ORIGINAL != "0") %>%
    pivot_longer(-PROV_KOD_ORIGINAL, names_to = "NRM_PARAMETERKOD", values_to = "ANALYS_DAT") %>%
    mutate(ANALYS_DAT = fix_date(ANALYS_DAT))
}

#' Title
#'
#' @param path
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file_weight <- function(path, sheet = 8){
  data <- readxl::read_excel(path = path, sheet = sheet)
  names(data)[1:6] <- c("PROV_KOD_ORIGINAL", "PROV_KOD_LABB", "GENUS", "PROVPLATS_ANALYSMALL", "DWEIGHT", "WWEIGHT")
  select(data, -PROV_KOD_LABB, -GENUS, -PROVPLATS_ANALYSMALL, -contains("...")) %>%
    filter(!(is.na(DWEIGHT) & is.na(WWEIGHT)))
}

#' Title
#'
#' @param path
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file_general <- function(path, sheet = "general info"){
  info <- readxl::read_excel(path = path, sheet = sheet)
  tibble(LABB = as.character(info[5, 2]),
         PROV_BERED = as.character(info[8, 2]),
         PROVKARL = as.character(info[9, 2]),
         ANALYS_MET = as.character(info[10, 2]),
         UTFOR_LABB = ifelse(str_detect(info[6, 2], "Click to choose"), "EJ_REL", as.character(info[6, 2])),
         ANALYS_INSTR = as.character(info[11, 2])) %>% 
    mutate(LABB = ifelse(str_detect(LABB, "ACES"), "ACES", LABB)) # ACES subdepartment should not be reported
}

#' Title
#'
#' @param path
#' @param sheet
#'
#' @return
#'
#'
#' @examples
read_lab_file_ackr <- function(path, sheet = "general info"){
  readxl::read_excel(path = path, sheet = sheet, skip = 27, col_names = c("NRM_PARAMETERKOD", "ACKREDITERAD_MET")) %>%
    filter(!is.na(NRM_PARAMETERKOD), !(ACKREDITERAD_MET == "…Click to choose…"))
}

#' Extracts information from lab protocol
#'
#' @param path
#' @param negative_for_nondetect
#' @param codes_path
#'
#' @return
#' @export
#'
#' @examples
moc_read_lab <- function(path, negative_for_nondetect = TRUE, codes_path = system.file("extdata", "codelist.xlsx", package = "MoCiS2")){
  suppressMessages({
    results <- read_lab_file(path)
    uncertainty <- read_lab_file2(path, "MATOSAKERHET", "uncertainty")
    LOD <- read_lab_file2(path, "DETEKTIONSGRANS_LOD", "LOD") %>% 
      mutate(DETEKTIONSGRANS_LOD = abs(DETEKTIONSGRANS_LOD))
    LOQ <- read_lab_file2(path, "RAPPORTERINGSGRANS_LOQ", "LOQ") %>% 
      mutate(RAPPORTERINGSGRANS_LOQ = abs(RAPPORTERINGSGRANS_LOQ))
    general <- read_lab_file_general(path)
    ackr <- read_lab_file_ackr(path)
    weight <- read_lab_file_weight(path)
    dates <- read_lab_file_date(path)
    koder_substans <- readxl::read_excel(codes_path, sheet = "PARAMETRAR")  %>%
      select(NRM_PARAMETERKOD, PARAMETERNAMN, UNIK_PARAMETERKOD, ENHET, MATOSAKERHET_ENHET, PROV_LAGR)
    koder_stationer <- select(results, PROVPLATS_ANALYSMALL) %>% distinct() %>%
      fuzzyjoin::stringdist_left_join(readxl::read_excel(codes_path, sheet = "STATIONER") %>%
                             select(-contains("...")) %>%
                             distinct(),
                           by = "PROVPLATS_ANALYSMALL", distance_col = "LOKDIST") %>%
      select(-PROVPLATS_ANALYSMALL.y, PROVPLATS_ANALYSMALL = PROVPLATS_ANALYSMALL.x) %>%
      distinct()
    koder_art <- select(results, LATIN = GENUS) %>% distinct() %>%
      fuzzyjoin::stringdist_left_join(readxl::read_excel(codes_path, sheet = "ARTER") %>%
                             select(-contains("...")) %>%
                             distinct(),
                           by = "LATIN", distance_col = "ARTDIST") %>%
      select(-LATIN.y, LATIN = LATIN.x) %>%
      distinct()
  })
  data <- left_join(results, uncertainty, by = c("PROV_KOD_ORIGINAL", "NRM_PARAMETERKOD")) %>%
    left_join(LOD, by = c("PROV_KOD_ORIGINAL", "NRM_PARAMETERKOD")) %>%
    left_join(LOQ, by = c("PROV_KOD_ORIGINAL", "NRM_PARAMETERKOD")) %>%
    left_join(dates, by = c("PROV_KOD_ORIGINAL", "NRM_PARAMETERKOD")) %>%
    left_join(weight, by = "PROV_KOD_ORIGINAL") %>%
    cbind(general) %>%
    left_join(koder_substans, by = "NRM_PARAMETERKOD") %>%
    left_join(ackr, by = "NRM_PARAMETERKOD") %>%
    mutate_at(c("PROV_BERED", "PROVKARL", "ANALYS_MET", "ANALYS_INSTR"), ~ifelse(str_detect(NRM_PARAMETERKOD, "FPRC|TPRC"), NA, .x)) %>%
    rename(LATIN = GENUS) %>%
    left_join(koder_art, by = "LATIN") %>%
    left_join(koder_stationer, by = "PROVPLATS_ANALYSMALL") %>%
    mutate(MATVARDETAL_ANM = if_else((abs(MATVARDETAL) <= RAPPORTERINGSGRANS_LOQ) |
                                      ((MATVARDETAL < 0) & (negative_for_nondetect)), "<", "", missing = ""),
           MATVARDETAL = ifelse(negative_for_nondetect, abs(MATVARDETAL), MATVARDETAL),
           MATV_STD = case_when(MATVARDETAL == DETEKTIONSGRANS_LOD ~ "b",
                                MATVARDETAL <= RAPPORTERINGSGRANS_LOQ ~ "q",
                                MATVARDETAL > RAPPORTERINGSGRANS_LOQ ~ "",
                                is.na(DETEKTIONSGRANS_LOD) & is.na(RAPPORTERINGSGRANS_LOQ) ~ ""),
           MATVARDESPAR = ifelse((MATVARDETAL <= RAPPORTERINGSGRANS_LOQ) & (MATVARDETAL > DETEKTIONSGRANS_LOD), "Ja", NA),
           MATOSAKERHET = ifelse(MATVARDETAL <= RAPPORTERINGSGRANS_LOQ, NA, MATOSAKERHET),
           MATOSAKERHET_ENHET = ifelse(is.na(MATOSAKERHET), NA, MATOSAKERHET_ENHET),
           ART = if_else((ART == "Stromming") & (LOC %in% c("VADO", "FLAD", "KULL", "ABBE", "HABU", "40G7", "UTLV", "UTLA")), "Sill", ART)) %>%
    filter(!is.na(MATVARDETAL)) %>%
    mutate_if(is.character, ~ifelse(str_detect(.x, "Click to choose"), NA, .x))
  if (max(data$ARTDIST, na.rm = TRUE) > 0){
    message("Warning: The following species were fuzzy matched")
    data %>% filter(ARTDIST > 0) %>% select(ART, LATIN) %>%
      mutate(match = paste("*", paste(LATIN, ART, sep = " -> "))) %>%
      pull(match) %>% unique() %>% paste(collapse = "\n") %>% message()
  }
  if (any(data$PROVPLATS_ANALYSMALL != data$NAMN_PROVPLATS)){
    message("Warning: The following locations were fuzzy matched")
    data %>% filter(PROVPLATS_ANALYSMALL != NAMN_PROVPLATS) %>%
      select(PROVPLATS_ANALYSMALL, NAMN_PROVPLATS) %>%
      mutate(match = paste("*", paste(PROVPLATS_ANALYSMALL, NAMN_PROVPLATS, sep = " -> "))) %>%
      pull(match) %>% unique() %>% paste(collapse = "\n") %>% message()
  }
  if (anyNA(data$ART)){
    message("WARNING: Unable to match the following species")
    data %>% filter(is.na(ART)) %>% pull(LATIN) %>% unique() %>% paste(collapse = ", ") %>% message()
  }
  if (anyNA(data$NAMN_PROVPLATS)){
    message("WARNING: Unable to match the following locations")
    data %>% filter(is.na(NAMN_PROVPLATS)) %>% pull(PROVPLATS_ANALYSMALL) %>% unique() %>% paste(collapse = ", ") %>% message()
  }
  if (anyNA(data$PARAMETERNAMN)){
    message("WARNING: Unable to match the following substance codes")
    data %>% filter(is.na(PARAMETERNAMN)) %>% pull(NRM_PARAMETERKOD) %>% unique() %>% paste(collapse = ", ") %>% message()
  }
  data %>% filter(!is.na(PARAMETERNAMN))
}

#' Title
#'
#' @param record
#'
#' @return
#' @export
#'
#' @examples
add_header_info <- function(record){
  header <- record$raw[1]
  # See code list for explanations
  mutate(record,
         YEAR = as.numeric(str_sub(header, 1, 4)),
         WEEK = as.numeric(str_sub(header, 5, 6)),
         DAY= as.numeric(str_sub(header, 7, 7)),
         GENUS = str_sub(header, 12, 15),
         LOC = str_sub(header, 46, 49),
         MYEAR = as.numeric(str_sub(header, 25, 26)),
         MYEAR = ifelse(MYEAR > 50, MYEAR + 1900, MYEAR + 2000),
         PROVTAG_DAT = as.Date(str_sub(header, 1, 7), "%Y%W%u")
  )
}


#' Title
#'
#' @param record
#'
#' @return
#' @export
#'
#' @examples
add_id_cols <- function(record){
  accnr <- filter(record, str_detect(NRM_CODE, "ACCNR")) %>% pull(VALUE)
  if (length(accnr) == 0)
    warning(paste("No ACCNR for record:", record$row[1]))
  if (length(accnr) == 0)
    warning(paste("Multiple ACCNR for record:", record$row[1], "picking first."))
  mutate(record,
         ACCNR = accnr[1],
         LAB_KOD = ifelse(Group == "lab", VALUE, NA)) %>%
    fill(LAB_KOD) %>%
    filter(!str_detect(NRM_CODE, "ACCNR"))
}

#' Title
#'
#' @param record
#'
#' @return
#' @export
#'
#' @examples
add_bio_cols <- function(record){
  mutate(record,
         NRM_CODE = ifelse(NRM_CODE == "KPRL", "KRPL", NRM_CODE), # Common misspelling
         #UNIT = ifelse(NRM_CODE %in% c("TOTV", "TOTL", "KRPL"), str_extract(VALUE, "[^,]+$") %>% tolower(), UNIT),
         ORGAN = ifelse(NRM_CODE == "ALDR", case_when(str_extract(VALUE, "[^,]+$") == "7" ~ "FJALL",
                                                      str_extract(VALUE, "[^,]+$") == "9" ~ "OTOLIT",
                                                      str_extract(VALUE, "[^,]+$") == "10" ~ "OPERCULUM"),
                        ORGAN),
         VALUE = ifelse(NRM_CODE %in% c("TOTV", "TOTL", "KRPL", "ALDR"), str_extract(VALUE, "[^,]*"), VALUE),
         TOTL = ifelse(NRM_CODE == "TOTL", as.numeric(VALUE), NA),
         KRPL = ifelse(NRM_CODE == "KRPL", as.numeric(VALUE), NA),
         TOTV = ifelse(NRM_CODE == "TOTV", as.numeric(VALUE), NA),
         ALDR = ifelse(NRM_CODE == "ALDR", as.numeric(VALUE), NA),
         ANM = ifelse(NRM_CODE %in% c("ANM", "NOTE"), VALUE, NA),
         SEX = ifelse(NRM_CODE == "SEX", as.numeric(VALUE), NA),
         NHOM = ifelse(str_length(ACCNR) == 17, as.numeric(str_sub(ACCNR, 13, 17)) - as.numeric(str_sub(ACCNR, 7, 11)) + 1, 1),
         NRM_CODE = ifelse((NRM_CODE %in% c("TOTV", "TOTL", "KRPL", "ALDR")) & str_detect(ACCNR, "-"), paste0(NRM_CODE, "H"), NRM_CODE)
  ) %>%
    fill(TOTL, KRPL, TOTV, ALDR, ANM, SEX, .direction = "downup")
}

#' Title
#'
#' @param record
#'
#' @return
#' @export
#'
#' @examples
add_prc_cols <- function(record){
  mutate(record,
         FPRC = ifelse(str_detect(NRM_CODE, "^FPRC"), as.numeric(VALUE), NA),
         LTPRC = ifelse(str_detect(NRM_CODE, "^LTPRC"), as.numeric(VALUE), NA),
         MTPRC = ifelse(str_detect(NRM_CODE, "^MTPRC"), as.numeric(VALUE), NA)) %>%
    group_by(LAB_KOD) %>%
    fill(FPRC, LTPRC, MTPRC, .direction = "downup") %>%
    ungroup()
}

#' Title
#'
#' @param record
#'
#' @return
#' @export
#'
#' @examples
frame_record <- function(record){
  record %>%
    fill(ORGAN, LAB) %>%
    add_header_info() %>%
    add_id_cols() %>%
    add_bio_cols() %>%
    add_prc_cols() %>%
    filter((label == "no")|is.na(label)) %>%
    select(-label)
}


#' Title
#'
#' @param prc_path
#' @param codes_path
#'
#' @return
#' @export
#'
#' @examples
moc_read_prc <- function(prc_path, codes_path = system.file("extdata", "codelist_prc.xls", package = "MoCiS2")){
  codes <- readxl::read_excel(codes_path) %>% select(Group, NRM_CODE, LAB, ORGAN, PARAMETER, label)
  tibble(raw = readLines(prc_path)) %>%
    mutate(NRM_CODE = str_extract(raw, "[^+]([^=]*)") %>% trimws() %>% toupper(),
           NRM_CODE = ifelse(str_sub(raw, 1, 1) %in% c("1", "2"), "HEADER", NRM_CODE),
           VALUE = str_extract(raw, "[^=]+$") %>% trimws()
    ) %>%
    left_join(codes, by = "NRM_CODE") %>%
    mutate(PARAMETER = ifelse(is.na(PARAMETER), NRM_CODE, PARAMETER),
           rec_id = cumsum(str_sub(raw, 1, 1) %in% c("1", "2"))) %>%
    group_by(rec_id) %>%
    nest() %>%
    pull(data) %>%
    map_df(frame_record) %>%
    mutate(VALUE = as.numeric(VALUE),
           VALUE = ifelse(VALUE %in% c(-9, -.9), NA, VALUE))
}
