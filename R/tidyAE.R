#' Read and Tidy Acoustic Emission Data
#'
#' @description
#' This function reads and processes Mistras Acoustic Emission (AE) data from an Excel file.
#'
#' @details
#'
#' A tibble is returned with:
#'
#' - Filtered to start from the row containing "ID"
#' - Filtered to remove rows where CH is not a number
#' - Date and time parsed from the original timestamp with microseconds processed
#' - Additional time-based columns added (DateDiff, WDay, Day, Hour)
#' - Filtered for a specific date range (2024-09-27 to 2024-10-15)
#' - Removed original timestamp column
#' - Reordered columns (character columns first, then numeric)
#'
#' @source
#' https://stackoverflow.com/questions/79157266/parsing-and-working-with-microsecond-precision-timestamps-in-r-using-dplyr
#'
#'
#' @param AEdata_filepath AEdata_filepath A string specifying the file path to the
#'          Excel file containing AE data.
#' @param xlsheet xlsheet A string specifying the name of the Excel sheet to read.
#'          Default is "Line Display".
#' @param skip_lines An integer specifying the number of lines to skip at the beginning of the Excel sheet.
#'          Default is 1.
#'
#' @return A tibble containing the processed AE data
#' @export
#'
#' @importFrom readxl read_excel
#' @importFrom dplyr filter mutate group_by lag select relocate
#' @importFrom lubridate dmy_hms wday floor_date hour as_date
#' @importFrom tools file_ext
#'
#' @examples
#'
#' \dontrun{
#' ae_data <- tidyAE("path/to/your/AEdata.xlsx")
#' }
#'
#'
tidyAE <- function(AEdata_filepath, xlsheet = "Line Display", skip_lines = 1) {

  #
  ext = tools::file_ext(AEdata_filepath)

  # Read xl file: use skip_lines if you know the header
  AEdatafile = readxl::read_excel(AEdata_filepath, sheet = xlsheet, skip = skip_lines)

  # Find start of data
  AEdata_start = which(grepl("ID", AEdatafile[[1]]))

  # Select data from the start
  AEdatafile = AEdatafile[AEdata_start:nrow(AEdatafile), ]

  ae_data <-
    AEdatafile |>
    dplyr::filter(!is.na(as.numeric(CH))) |>
    dplyr::group_by(CH) |>
    dplyr::mutate(
      Date = dmy_hms(sub("( \\d{2})([ap])(.*$)", "\\1\\3 \\2m", `DD/MM/YY  HH:MM:SS.mmmuuun`)),
      ms = format(Date, "%OS6"),
      DateDiff = Date - lag(Date),
      WDay = lubridate::wday(Date, label = TRUE),
      Day = lubridate::floor_date(Date, unit = "day"),
      Hour = lubridate::hour(Date)
    ) |>
    dplyr::filter(
      Date > lubridate::as_date("2024-09-27"),
      Date < lubridate::as_date("2024-10-15"),
      ENER > 0) |>
    dplyr::select(-`DD/MM/YY  HH:MM:SS.mmmuuun`) |>
    dplyr::relocate(where(is.numeric), .after = where(is.character))

  return(ae_data)
}
