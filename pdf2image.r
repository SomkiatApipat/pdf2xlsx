

library(pdftools)

pdf2xlsx <- function(input_pdf) {

  # Get the number of pages in the PDF file
  num_pages <- pdf_info(input_pdf)$pages
  pages = 1:num_pages

  # get file name only
  # Split the file name into parts using the "." as a delimiter
  file_parts <- strsplit(input_pdf, "\\.")
  file_name_only <- file_parts[[1]][1]

  # assign
  file_type <- "png"
  res_pdi <- 200

  for (i in 1:num_pages) {
    output_png <- paste0(file_name_only,"-page",pages[i],".",file_type)
    #pdf_convert(input_pdf, format = file_type, pages = pages[i],
    #           filenames = output_png, dpi = resolution_pdi)

    bitmap <- pdf_render_page(input_pdf, page = i, dpi = res_pdi)
    png::writePNG(bitmap, output_png)
  }





}

