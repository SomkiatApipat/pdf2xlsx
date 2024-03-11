#' Convert pdf (data table) to xlsx (excel)
#'
#' Getting all allocated projects in a given province
#'
#' @param input_pdf Character, a name of pdf file
#'
#' @return xlsx file contains all allocated projects
#' @export
#'
#' @examples input_pdf <- system.file("extdata","0150000s.pdf",package = "pdf2xlsx")
#' @examples browseURL(input_pdf)
#' @examples data <- pdf2xlsx(input_pdf)
#' @examples View(data)
#'
#' @import pdftools
#' @import tesseract
#' @import magick
#' @import png
#' @import readr
#'

pdf2xlsx <- function(input_pdf) {

  # Only need to do download once:
  tesseract_download("tha")
  # Now load the dictionary
  (thai <- tesseract("tha"))

  # Get the number of pages in the PDF file
  num_pages <- pdf_info(input_pdf)$pages
  pages = 1:num_pages

  # get file name only
  # Split the file name into parts using the "." as a delimiter
  file_parts <- strsplit(input_pdf, "\\.")
  file_name_only <- file_parts[[1]][1]

  ### --- convert pdf to png --- ###

  # assigned png property
  file_type <- "png"
  res_pdi <- 400

  for (i in 1:num_pages) {

    output_png <- paste0(file_name_only,"-page",pages[i],".",file_type)
    #pdf_convert(input_pdf, format = file_type, pages = pages[i],
    #           filenames = output_png, dpi = res_pdi)

    bitmap <- pdf_render_page(input_pdf, page = pages[i], dpi = res_pdi)
    png::writePNG(bitmap, output_png)
  }

  ### --- prepare table image png --- ###
  for (i in 1:num_pages) {

    input_png <- paste0(file_name_only,"-page",pages[i],".",file_type)

    #############################################################
    # find corners of table image and crop the table

    # using png package to find corners of table image
    # read image into bitmap format
    img <- readPNG(readBin(input_png, "raw", 1e6))

    #m(img)
    #Plot(as.raster(img))

    # top and bottom horizontal lines
    sum_horiz <- apply(img[,,2], 1, sum)
    hdiff = max(sum_horiz) - min(sum_horiz)
    hvalue = min(sum_horiz) + 0.2*hdiff # adding
    #plot(sum_horiz, type = "l")
    id = 1:length(sum_horiz)
    target_row <- id[sum_horiz<=hvalue]
    top_hline = min(target_row)
    bot_hline = max(target_row)

    # most left and most right vertical lines
    sum_vert <- apply(img[,,2], 2, sum)
    vdiff = max(sum_vert) - min(sum_vert)
    vvalue = min(sum_vert) + 0.2*vdiff
    id = 1:length(sum_vert)
    # plot(sum_vert, type = "l")
    target_colm <- id[sum_vert<vvalue]
    left_vline = min(target_colm)
    right_vline = max(target_colm)

    # using magick package to crop table image
    # read image file
    raw_img <- image_read(input_png)
    # Image details
    info <- image_info(raw_img)
    # You can also access to each element
    pg_width = info$width # find image width
    pg_height = info$height # find image height

    # using values of right_vline, left_vline,
    #  bot_hline, and top_vling to crop table image
    crop_width = right_vline - left_vline
    crop_height = bot_hline - top_hline

    set_bound = paste(crop_width,"x",crop_height,"+",left_vline,"+",top_hline,"")

    img_crop = image_crop(raw_img, geometry = set_bound) # save raw img
    #image_write(img_crop,"test-crop.png")

    ####################################################
    # change any background colors to white color and
    # remove grid lines
    # using magick package

    ##############################################################
    # remove the grid by converting “white” colors to transparent,
    # and allowing for some “fuzzy” approximation of colors that are close to white
    # or “touching”. There’s a lot more to fuzz in that it’s the “relative color
    # distance (value between 0 and 100) to be considered similar in the filling
    # algorithm”, but I’m not a color space expert by any means.
    bg_white <- img_crop %>%
      image_transparent(color = "white", fuzz=50) %>%
      image_background("white")

    # remove grid lines in image
    no_grid=image_fill(bg_white, "white", point = "+1+1")

    #image_ggplot(no_grid)
    #filename = "page1_bg_white.png"
    #image_write(no_grid, filename_out)
    #print(paste("white background image is created and named:",filename_out))


    ### --- convert png to data frame --- ###

    ###################################################
    # read text from final image
    # read final image
    text <- ocr(no_grid, engine = thai)
    #cat(text)

    # setting text lines
    zz = read_lines(text)
    nx = nchar(zz)
    #zz = zz[nx>50]  # remove blank lines
    zz = zz[nx>0]  # remove blank lines
    zz = zz[-(1)] # remove heading

    # separate columns and put them in data from
    nlines=length(zz)
    nx = nchar(zz)
    dataf = data.frame (
      item = rep("",nlines),
      unit = rep("",nlines),
      amount = rep("",nlines),
      budget = rep("",nlines)
    )
      for (j in 1:nlines) {
        non_white_space_positions <- unlist(gregexpr("\\S", zz[j]))
        diff_positions = diff(non_white_space_positions)
        pos = 1:length(diff_positions)
        indx_positions = non_white_space_positions[pos[diff_positions>5]]
        lindx = length(indx_positions)
        if (lindx == 0) {
          temp = substr(zz[j],start=1,stop=nx[j])
          temp <- stringr::str_replace_all(temp,"[.]", "")
          temp <- stringr::str_replace_all(temp,"[,]", "")
          temp <- stringr::str_replace_all(temp,"[ ]", "")
          if (is.na(as.numeric(temp)))  {
            dataf$item[j] = temp
          } else {
            dataf$budget[j] = temp
          }

      } else if (lindx == 1) {

          temp = substr(zz[j],start=1,stop=indx_positions[1])
          if (is.na(as.numeric(temp)))  {
            dataf$item[j] = temp
          } else {
            dataf$amount[j] = temp
          }
          dataf$budget[j] = substr(zz[j],start=indx_positions[1]+1,stop=nx[j])

      } else if (lindx == 2) {

          temp = substr(zz[j],start=1,stop=indx_positions[1])
          if (nchar(temp)<8)  {
            dataf$unit[j] = temp
          } else {
            dataf$item[j] = temp
          }
          dataf$amount[j] = substr(zz[j],start=indx_positions[1]+1,
                                   stop=indx_positions[2])
          dataf$budget[j] = substr(zz[j],start=indx_positions[2]+1,stop=nx[j])

        } else if (lindx == 3) {
          dataf$item[j] = substr(zz[j],start=1,stop=indx_positions[1])
          dataf$unit[j] = substr(zz[j],start=indx_positions[1]+1,
                                 stop=indx_positions[2])
          dataf$amount[j] = substr(zz[j],start=indx_positions[2]+1,
                                   stop=indx_positions[3])
          dataf$budget[j] = substr(zz[j],start=indx_positions[3]+1,stop=nx[j])
        } else {
          #  print (" error in separate colm; lindx > 3")
        }
    }

    # remove white space before and after text
    dataf$item=trimws(dataf$item)
    dataf$unit=trimws(dataf$unit)
    dataf$amount=trimws(dataf$amount)
    dataf$budget=trimws(dataf$budget)

    # remove rows that dataf$item shorter than 5
    #id = 1:nrow(dataf)
    #redund = id[nchar(dataf$item)<4 | nchar(dataf$unit)==1]
    #if (length(redund) > 0) {dataf = dataf[-redund,]}

    # clean up number and as.numeric
    #dataf$amount <- stringr::str_replace_all(dataf$amount,"[.]", "")
    dataf$amount <- stringr::str_replace_all(dataf$amount,"[,]", "")

    dataf$budget <- stringr::str_replace_all(dataf$budget,"[.]", "")
    dataf$budget <- stringr::str_replace_all(dataf$budget,"[,]", "")
    dataf$budget <- stringr::str_replace_all(dataf$budget,"[ ]", "")

    # delete dataf that dataf$amount with some character but connot convert to numeric
    id = 1:nrow(dataf)
    redund = id[(nchar(dataf$amount)>0 & is.na(as.numeric(dataf$amount))) |
                  (nchar(dataf$budget)>0 & is.na(as.numeric(dataf$budget)))] #|
    #            nchar(dataf$item)<4 | nchar(dataf$unit)==1 ]
    if (length(redund) > 0) {dataf = dataf[-redund,]}

    dataf$amount <- as.numeric(dataf$amount)
    dataf$budget <- as.numeric(dataf$budget)

    # delete dataf$budget < 100
    #id = 1:nrow(dataf)
    #redund = id[!is.na(dataf$budget) & dataf$budget < 100]
    #if (length(redund) > 0) {dataf = dataf[-redund,]}

    # adding page number
    ndata = nrow(dataf)
    pp = rep(pages[i],ndata)
    dataf <- cbind(dataf,pp)

    if (i == 1) {
      dataf_final = dataf
    } else {
      dataf_final = rbind(dataf_final,dataf)
    }

    cat("page",i,"out of",num_pages,"...done","\n")

  }

  ### --- combine items --- ###
  # delete blank lines
  id = 1:nrow(dataf_final)
  line_blank = id[nchar(dataf_final$item)==0 & is.na(dataf_final$budget)]
  if (length(line_blank)>0) {dataf_final = dataf_final[-line_blank,]}

  # manage df lines that have no item
  id = 1:nrow(dataf_final)
  line_no_item = id[nchar(dataf_final$item)==0 & !is.na(dataf_final$budget)]
  num_line = length(line_no_item)

  ### copy lines to the above lines
  if (num_line > 0) {

    for (k in 1:num_line) {
      ll = line_no_item[k]
      lb = ll-1
      dataf_final$unit[lb] = dataf_final$unit[ll]
      dataf_final$amount[lb] = dataf_final$amount[ll]
      dataf_final$budget[lb] = dataf_final$budget[ll]
    }
    dataf_final = dataf_final[-line_no_item,] # remove lines that moved up
  }


  # combine lines
  id = 1:nrow(dataf_final)
  item_no_budget = id[is.na(dataf_final$budget)]
  num_item = length(item_no_budget)

  diff_item = diff(item_no_budget)
  ndiff = length(diff_item)
  st = rep(0,ndiff)
  ed = rep(0,ndiff)

  for (i in 1:ndiff) {
    if (diff_item[i]==1) {
      st[i] = item_no_budget[i]
      ed[i] = item_no_budget[i]+2

      if (i!=1) {
        if (diff_item[i-1]==1) {ed[i-1]=ed[i]}
      } else {
        ed[i] = item_no_budget[i]+2
      }

    } else {
      st[i] = item_no_budget[i]
      ed[i] = item_no_budget[i]+1
    }
  }
  st[ndiff+1] = item_no_budget[num_item]
  ed[ndiff+1] = st[ndiff+1]+1

  duplicated_ed <- unique(ed[duplicated(ed)])
  ndup = length(duplicated_ed)
  for (i in 1:ndup) {
    st[ed==duplicated_ed[i]] = min(st[ed==duplicated_ed[i]])
  }

  # set start and end of items
  st_final = unique(st)
  ed_final = unique(ed)
  ncomb_item = length(st_final)
  comb_item = rep("",ncomb_item)

  # combine items
  for (i in 1:ncomb_item) {
    temp = ""
    for (j in st_final[i]:ed_final[i]) {
      temp = paste(temp,dataf_final$item[j])
    }
    comb_item[i] = temp
  }

  # replace combined items to end lines
  for (i in 1:ncomb_item) {
    dataf_final$item[ed_final[i]] = comb_item[i]
  }

  # delete item lines without budget
  dataf_final = dataf_final[-item_no_budget,]

  # remove white space in front and back of items
  dataf_final$item = trimws(dataf_final$item)

  return(dataf_final)

}

