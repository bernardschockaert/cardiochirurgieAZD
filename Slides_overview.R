# =============================================================================
#  Cardiochirurgie AZ Delta -- two-slide PowerPoint overview
# -----------------------------------------------------------------------------
#  Builds output/Cardiochirurgie_overview.pptx from the derived dataset:
#    slide 1  the cohort and its baseline
#    slide 2  early recovery and outcomes
#
#  Everything is drawn as NATIVE PowerPoint shapes (DrawingML), so every
#  rectangle, line and word stays selectable and editable in PowerPoint -
#  no pictures, no fonts baked into an image, file size a few tens of kB.
#
#  Requires: officer   (install.packages("officer")).  Nothing else - no
#            ggplot2, no rvg.
#
#  HOW TO USE
#    source("Cardiochirurgie.R")   # once, to build the derived dataset
#    source("Slides_overview.R")
#  Slides_overview.R re-uses the object 'd' if it is already in the session,
#  otherwise it reads output/analysis_dataset.rds, otherwise it sources
#  Cardiochirurgie.R for you.
# =============================================================================


# =============================================================================
# 0.  CONFIGURATION
# =============================================================================

PPTX_OUT   <- file.path("output", "Cardiochirurgie_overview.pptx")
SLIDE_W    <- 13.333          # inches, 16:9 widescreen
SLIDE_H    <- 7.5

## Font. "Segoe UI" ships with Windows Office; on macOS use "Helvetica Neue"
## or "Calibri". Whatever you pick must exist on the machine that opens the
## deck, otherwise PowerPoint substitutes it.
FONT <- "Segoe UI"

## Palette, sampled from the example report ----------------------------------
NAVY   <- "2F4D5D"   # hero band
SLATE  <- "37596B"   # primary bars / chart ink
SLATE2 <- "6E8C9C"   # secondary bars
SLATE3 <- "A9C0CB"   # tertiary bars
LIME   <- "C6E86B"   # accent for highlighted figures
INK    <- "17323F"   # body text
MUTED  <- "6F8794"   # secondary text
FAINT  <- "9DB2BD"   # text on the navy band
PANEL  <- "EEF4F8"   # card fill
RULE   <- "D8E3EC"   # hairlines and card borders
WHITE  <- "FFFFFF"

## Dutch number formatting (1.892 and 27,4) ----------------------------------
nl_int <- function(x) formatC(round(x), format = "d", big.mark = ".", decimal.mark = ",")
nl_num <- function(x, d = 1) sub("\\.", ",", formatC(round(x, d), format = "f", digits = d))
nl_pct <- function(x, d = 1) paste0(nl_num(x, d), " %")


# =============================================================================
# 1.  DRAWINGML SHAPE LAYER
#     Small helpers that emit raw PowerPoint shapes. officer inserts them with
#     ph_with(); PowerPoint then treats them as ordinary shapes.
# =============================================================================

if (!requireNamespace("officer", quietly = TRUE))
  stop("Package 'officer' is required.  install.packages('officer')")

EMU <- 914400L                                   # English Metric Units per inch
e   <- function(inch) sprintf("%.0f", inch * EMU)

.id <- local({ i <- 100L; function() { i <<- i + 1L; i } })

## XML-escape, and turn every non-ASCII character into a numeric character
## reference (e.g. "e" with diaeresis -> "&#235;").  The generated XML is then
## pure ASCII, so it parses correctly whatever locale R happens to run in -
## on a Windows RStudio with a non-UTF-8 locale this is what stops accented
## Dutch text from corrupting the file.
esc <- function(s) {
  s <- enc2utf8(as.character(s))
  s <- gsub("&", "&amp;", s, fixed = TRUE)
  s <- gsub("<", "&lt;",  s, fixed = TRUE)
  s <- gsub(">", "&gt;",  s, fixed = TRUE)
  vapply(s, function(one) {
    cp <- tryCatch(utf8ToInt(one), error = function(e) NA_integer_)
    if (anyNA(cp)) return(iconv(one, to = "ASCII", sub = "?"))
    if (!any(cp > 127)) return(one)
    paste(ifelse(cp > 127, paste0("&#", cp, ";"), intToUtf8(cp, multiple = TRUE)),
          collapse = "")
  }, character(1), USE.NAMES = FALSE)
}

NS <- paste('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')

## One styled piece of text.
##   size in pt, spc = letter spacing in pt (positive opens the tracking out)
run <- function(text, size = 11, color = INK, bold = FALSE, italic = FALSE,
                spc = 0, caps = FALSE, font = FONT)
  list(text = text, size = size, color = color, bold = bold, italic = italic,
       spc = spc, caps = caps, font = font)

## One paragraph = a list of runs plus paragraph-level settings.
para <- function(..., align = "l", space_before = 0, line = 100) {
  runs <- list(...)
  if (length(runs) == 1 && is.list(runs[[1]]) && is.null(runs[[1]]$text)) runs <- runs[[1]]
  list(runs = runs, align = align, space_before = space_before, line = line)
}

.run_xml <- function(r)
  sprintf('<a:r><a:rPr lang="nl-BE" sz="%d" b="%d" i="%d" spc="%d"%s dirty="0"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:latin typeface="%s"/><a:cs typeface="%s"/></a:rPr><a:t>%s</a:t></a:r>',
          round(r$size * 100), as.integer(r$bold), as.integer(r$italic),
          round(r$spc * 100), if (r$caps) ' cap="all"' else "",
          r$color, r$font, r$font, esc(r$text))

.para_xml <- function(p)
  sprintf('<a:p><a:pPr algn="%s" marL="0" indent="0"><a:lnSpc><a:spcPct val="%d"/></a:lnSpc><a:spcBef><a:spcPts val="%d"/></a:spcBef><a:buNone/></a:pPr>%s</a:p>',
          switch(p$align, l = "l", c = "ctr", r = "r", "l"), round(p$line * 1000),
          round(p$space_before * 100),
          if (length(p$runs)) paste(vapply(p$runs, .run_xml, ""), collapse = "") else "")

.fill_xml <- function(fill) {
  if (is.null(fill) || is.na(fill)) return("<a:noFill/>")
  sprintf('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>', fill)
}

.line_xml <- function(line, lwd, dash = NULL) {
  if (is.null(line) || is.na(line)) return("<a:ln><a:noFill/></a:ln>")
  sprintf('<a:ln w="%.0f" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>',
          lwd * 12700, line,
          if (is.null(dash)) "" else sprintf('<a:prstDash val="%s"/>', dash))
}

## A rectangle (radius > 0 gives a rounded rectangle), optionally with text.
sp_rect <- function(l, t, w, h, fill = NA, line = NA, lwd = 1, radius = 0,
                    paras = NULL, anchor = "ctr", pad = 0.10) {
  geom <- if (radius > 0)
    sprintf('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val %.0f"/></a:avLst></a:prstGeom>',
            min(50000, radius / min(w, h) * 100000))
  else '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  body <- if (is.null(paras)) '<a:p/>' else paste(vapply(paras, .para_xml, ""), collapse = "")
  xml <- sprintf('<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="s%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>%s%s%s</p:spPr>
<p:txBody><a:bodyPr anchor="%s" lIns="%s" rIns="%s" tIns="%s" bIns="%s" wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>',
    NS, .id(), .id(), e(l), e(t), e(w), e(h), geom, .fill_xml(fill), .line_xml(line, lwd),
    anchor, e(pad), e(pad), e(pad * 0.4), e(pad * 0.4), body)
  list(xml = xml, l = l, t = t, w = w, h = h)
}

## A text box with no fill and no border.
sp_txt <- function(l, t, w, h, paras, anchor = "t", pad = 0)
  sp_rect(l, t, w, h, fill = NA, line = NA, paras = paras, anchor = anchor, pad = pad)

## A straight line between two points.
sp_line <- function(x1, y1, x2, y2, color = RULE, lwd = 1, dash = NULL) {
  flipH <- x2 < x1; flipV <- y2 < y1
  l <- min(x1, x2); t <- min(y1, y2)
  w <- abs(x2 - x1); h <- abs(y2 - y1)
  xml <- sprintf('<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="l%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm%s%s><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>%s</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>',
    NS, .id(), .id(),
    if (flipH) ' flipH="1"' else "", if (flipV) ' flipV="1"' else "",
    e(l), e(t), e(w), e(h), .line_xml(color, lwd, dash))
  list(xml = xml, l = l, t = t, w = w, h = h)
}

## A circle / ellipse.
sp_ellipse <- function(l, t, w, h, fill = NA, line = NA, lwd = 1, paras = NULL) {
  xml <- sprintf('<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="e%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>
<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>%s%s</p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>%s</p:txBody></p:sp>',
    NS, .id(), .id(), e(l), e(t), e(w), e(h), .fill_xml(fill), .line_xml(line, lwd),
    if (is.null(paras)) "<a:p/>" else paste(vapply(paras, .para_xml, ""), collapse = ""))
  list(xml = xml, l = l, t = t, w = w, h = h)
}

## Push a list of shapes onto a slide.
## officer replaces a shape's own <a:xfrm> with the ph_location it is given,
## so each shape carries its geometry and we hand the same values back.
add_shapes <- function(doc, shapes) {
  for (s in shapes)
    doc <- officer::ph_with(
      doc,
      value    = xml2::read_xml(s$xml),
      location = officer::ph_location(left = s$l, top = s$t,
                                      width = max(s$w, 0.01), height = max(s$h, 0.01)))
  doc
}


# =============================================================================
# 2.  COMPOSITE ELEMENTS
#     The building blocks the two slides are assembled from.
# =============================================================================

## Section heading: short accent rule + title.
el_heading <- function(l, t, title, w = 4, color = INK, rule = SLATE, size = 12.5)
  list(sp_line(l, t + 0.115, l + 0.22, t + 0.115, rule, 2.25),
       sp_txt(l + 0.32, t - 0.06, w, 0.34,
              list(para(run(title, size, color, bold = TRUE)))))

## Horizontal bar chart: one row per category, value shown at the right.
## `frac` scales the bar; `label_w` is the width of the category column.
el_hbars <- function(l, t, w, values, labels, fills, label_w = 1.62,
                     row_h = 0.315, gap = 0.075, max_val = NULL,
                     val_fmt = function(v) nl_pct(v), track = TRUE) {
  mx <- if (is.null(max_val)) max(values, na.rm = TRUE) else max_val
  val_w   <- 0.62
  track_w <- w - label_w - val_w - 0.12
  out <- list()
  for (i in seq_along(values)) {
    y <- t + (i - 1) * (row_h + gap)
    bh <- 0.145                                  # bar thickness
    by <- y + (row_h - bh) / 2
    out <- c(out, list(
      sp_txt(l, y, label_w - 0.08, row_h,
             list(para(run(labels[i], 9.5, INK))), anchor = "ctr")))
    if (track)
      out <- c(out, list(sp_rect(l + label_w, by, track_w, bh, fill = PANEL, radius = bh / 2)))
    out <- c(out, list(sp_rect(l + label_w, by, max(0.03, track_w * values[i] / mx), bh,
                               fill = fills[i], radius = bh / 2)))
    out <- c(out, list(
      sp_txt(l + label_w + track_w + 0.10, y, val_w, row_h,
             list(para(run(val_fmt(values[i]), 9.5, INK, bold = TRUE), align = "r")),
             anchor = "ctr")))
  }
  out
}

## Vertical bar chart with a value on top and a category underneath.
el_vbars <- function(l, t, w, h, values, labels, fills,
                     val_fmt = function(v) nl_pct(v, 0)) {
  n  <- length(values)
  gp <- 0.13
  bw <- (w - gp * (n - 1)) / n
  mx <- max(values, na.rm = TRUE)
  plot_h <- h - 0.46                              # leave room for the axis label
  out <- list()
  for (i in seq_len(n)) {
    x  <- l + (i - 1) * (bw + gp)
    bh <- max(0.06, plot_h * values[i] / mx * 0.82)
    out <- c(out, list(
      sp_rect(x, t + plot_h - bh, bw, bh, fill = fills[i]),
      sp_txt(x, t + plot_h - bh - 0.28, bw, 0.26,
             list(para(run(val_fmt(values[i]), 9.5, INK, bold = TRUE), align = "c"))),
      sp_txt(x, t + plot_h + 0.06, bw, 0.24,
             list(para(run(labels[i], 8.5, MUTED), align = "c")))))
  }
  out
}

## Key/value table with hairline separators; `accent` marks rows to highlight.
el_table <- function(l, t, w, keys, vals, row_h = 0.315, accent = integer(0),
                     key_size = 9.5, val_size = 9.5) {
  out <- list()
  for (i in seq_along(keys)) {
    y  <- t + (i - 1) * row_h
    kc <- if (i %in% accent) SLATE else INK
    out <- c(out, list(
      sp_txt(l, y, w * 0.68, row_h, list(para(run(keys[i], key_size, kc))), anchor = "ctr"),
      sp_txt(l, y, w, row_h,
             list(para(run(vals[i], val_size, INK, bold = TRUE), align = "r")), anchor = "ctr"),
      sp_line(l, y + row_h, l + w, y + row_h, RULE, 0.75)))
  }
  out
}

## One KPI cell for the strip in the hero band.
el_kpi <- function(l, t, w, value, label, value_color = WHITE, size = 21)
  list(sp_txt(l, t, w, 0.42, list(para(run(value, size, value_color, bold = TRUE)))),
       sp_txt(l, t + 0.44, w, 0.26,
              list(para(run(label, 7.6, FAINT, caps = TRUE, spc = 0.9)))))

## Step line for a cumulative distribution, drawn as real line segments.
el_step <- function(l, t, w, h, x, y, color = SLATE, lwd = 2.25, dash = NULL) {
  sx <- function(v) l + w * (v - min(x)) / (max(x) - min(x))
  sy <- function(v) t + h * (1 - v / 100)
  out <- list()
  for (i in seq_len(length(x) - 1)) {
    out <- c(out, list(
      sp_line(sx(x[i]),     sy(y[i]), sx(x[i + 1]), sy(y[i]),     color, lwd, dash),
      sp_line(sx(x[i + 1]), sy(y[i]), sx(x[i + 1]), sy(y[i + 1]), color, lwd, dash)))
  }
  out
}


# =============================================================================
# 3.  DATA
# =============================================================================

if (!exists("d") || !is.data.frame(d)) {
  rds <- file.path("output", "analysis_dataset.rds")
  if (file.exists(rds)) {
    d <- readRDS(rds)
    message("Slides: using ", rds)
  } else {
    message("Slides: no derived dataset found, running Cardiochirurgie.R first ...")
    source("Cardiochirurgie.R")
  }
}

pc  <- function(x) 100 * mean(x, na.rm = TRUE)
med <- function(x) median(x, na.rm = TRUE)
iqr <- function(x, dg = 1) sprintf("%s-%s", nl_num(quantile(x, .25, na.rm = TRUE), dg),
                                            nl_num(quantile(x, .75, na.rm = TRUE), dg))

K <- list(
  n_rec      = nrow(d),
  n_pat      = length(unique(d$pat_id)),
  period     = sprintf("%s - %s", format(min(d$surg_start), "%m/%Y"),
                                  format(max(d$surg_start), "%m/%Y")),
  age_med    = med(d$age_years),
  age_iqr    = iqr(d$age_years, 0),
  male       = pc(d$sex == "Male"),
  elective   = pc(d$elective_f == "Elective"),
  bmi        = med(d$bmi),
  smoker     = pc(d$current_smoker),
  dm         = pc(d$diabetes_any),
  htn        = pc(d$hypertension),
  cva        = pc(d$prior_cva_tia),
  lung       = pc(d$lung_disease),
  mort_icu   = pc(d$death_icu),
  mort_30    = pc(d$death_30d),
  mort_180   = pc(d$death_180d),
  aki_any    = pc(d$aki_any),
  extub_6h   = pc(!d$vent_gt6h),
  on_table   = pc(d$extub_status == "Extubated on table"),
  vent_med   = med(d$vent_min_recorded[d$extub_status == "Ventilated on ICU"]) / 60,
  readm      = pc(d$icu_readmission),
  los_med    = med(d$icu_los_days),
  los_iqr    = iqr(d$icu_los_days, 1),
  chair_med  = med(d$time_to_chair_h),
  chair_24   = pc(d$chair_within_24h),
  vas_med    = med(d$mean_vas_24h)
)

age_band <- cut(d$age_years, c(-Inf, 60, 70, 80, Inf), right = FALSE,
                labels = c("< 60", "60-69", "70-79", "\u2265 80"))
K$age_dist  <- as.numeric(100 * prop.table(table(age_band)))
K$surg_dist <- as.numeric(100 * prop.table(table(d$surg_group)))
K$aki_dist  <- as.numeric(100 * prop.table(table(d$aki_stage)))

## cumulative share extubated, by hours since end of surgery
ext_h <- c(0, 1, 2, 3, 4, 6, 8, 12, 24)
K$ext_cum <- vapply(ext_h, function(hh) pc(d$time_to_extub_min <= hh * 60), 0)


# =============================================================================
# 4.  CREATE THE DECK AND RESOLVE THE PAGE GEOMETRY
# =============================================================================

## officer 0.6.x cannot set the slide size, and its default template is 4:3.
## Rewrite <p:sldSz> inside the package so the deck is truly 16:9, then read the
## size back and derive the layout from whatever we actually got - if the
## resize is ever refused, the slides simply lay themselves out on the
## template's own page instead of running off the edge.
resize_pptx <- function(path, width, height) {
  tmp <- file.path(tempdir(), sprintf("pptx_resize_%s", as.integer(Sys.time())))
  unlink(tmp, recursive = TRUE); dir.create(tmp, recursive = TRUE)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  utils::unzip(path, exdir = tmp)
  f <- file.path(tmp, "ppt", "presentation.xml")
  if (!file.exists(f)) stop("presentation.xml not found")
  x  <- paste(readLines(f, warn = FALSE), collapse = "")
  x2 <- sub("<p:sldSz[^>]*/>",
            sprintf('<p:sldSz cx="%.0f" cy="%.0f"/>', width * EMU, height * EMU), x)
  if (identical(x, x2)) stop("<p:sldSz> not found")
  con <- file(f, open = "wb"); writeBin(charToRaw(x2), con); close(con)
  ## zipr() keeps the directory structure; zip(mode = "cherry-pick") flattens
  ## it and produces a package PowerPoint cannot open.
  zip::zipr(zipfile = path,
            files = list.files(tmp, full.names = TRUE, all.files = TRUE, no.. = TRUE))
  invisible(TRUE)
}

new_deck <- function(width, height) {
  tmp <- tempfile(fileext = ".pptx")
  print(officer::read_pptx(), target = tmp)
  ok <- tryCatch({ resize_pptx(tmp, width, height); TRUE },
                 error = function(e) { warning("Could not set the slide size to ",
                   width, " x ", height, " in (", conditionMessage(e),
                   "); using the template page size instead.", call. = FALSE); FALSE })
  officer::read_pptx(tmp)
}

doc <- new_deck(SLIDE_W, SLIDE_H)
sz  <- officer::slide_size(doc)
SLIDE_W <- sz$width; SLIDE_H <- sz$height          # the size we really have
cat(sprintf("Slides: page %.3f x %.3f in\n", SLIDE_W, SLIDE_H))

## --- layout grid, derived from the page so it adapts to any slide size ------
M      <- 0.62                                     # page margin
GAP    <- 0.34                                     # gutter between columns
HERO_H <- SLIDE_H * 0.403                          # navy band height
CW     <- (SLIDE_W - 2 * M - 2 * GAP) / 3          # body column width
C1 <- M; C2 <- M + CW + GAP; C3 <- M + 2 * (CW + GAP)
C3W <- SLIDE_W - M - C3
NUMX <- SLIDE_W - M - 3.9                          # left edge of the hero figure


# =============================================================================
# 5.  SLIDE 1 -- THE COHORT
# =============================================================================

slide1 <- c(
  ## ---- hero band ----------------------------------------------------------
  list(sp_rect(0, 0, SLIDE_W, HERO_H, fill = NAVY)),
  list(sp_txt(M, 0.36, NUMX - M - 0.2, 0.28, list(para(
    run("CARDIOCHIRURGIE AZ DELTA", 8.4, FAINT, bold = TRUE, caps = TRUE, spc = 1.8),
    run("   \u00b7   CARDIO-ANESTHESIE & INTENSIEVE ZORG", 8.4, FAINT, caps = TRUE, spc = 1.8))))),
  list(sp_txt(M, 0.62, 8.0, 0.78, list(para(run("Cardiochirurgie & IZ", 34, WHITE, bold = TRUE))))),
  list(sp_txt(M, 1.46, NUMX - M - 0.3, 0.66, list(
    para(run(paste("Elke hartingreep steunt op een tweede team. Cardio-anesthesie en cardiale intensieve",
                   "zorg bewaken de pati\u00ebnt van inductie tot extubatie: hemodynamische sturing,",
                   "vroege extubatie en snelle mobilisatie volgens ERAS-principes."), 10.5, FAINT), line = 132)))),

  ## headline figure, right of the hero
  list(sp_txt(NUMX, 0.52, SLIDE_W - M - NUMX, 1.02,
              list(para(run(nl_int(K$n_rec), 50, WHITE, bold = TRUE), align = "r")))),
  list(sp_txt(NUMX, 1.56, SLIDE_W - M - NUMX, 0.26, list(para(
    run("IZ-OPNAMES NA HARTCHIRURGIE", 8.4, LIME, bold = TRUE, caps = TRUE, spc = 1.5), align = "r")))),
  list(sp_txt(NUMX, 1.82, SLIDE_W - M - NUMX, 0.26, list(para(
    run(sprintf("%s pati\u00ebnten \u00b7 %s", nl_int(K$n_pat), K$period), 9, FAINT), align = "r")))),

  ## ---- KPI strip ----------------------------------------------------------
  list(sp_line(M, 2.24, SLIDE_W - M, 2.24, "4A6B7C", 0.75))
)

kpis <- list(
  list(nl_num(K$age_med, 0),               "MEDIANE LEEFTIJD (JAAR)", WHITE),
  list(nl_pct(K$male, 1),                  "MANNEN",                  WHITE),
  list(nl_pct(K$elective, 1),              "ELECTIEVE INGREPEN",      LIME),
  list(nl_num(K$bmi, 1),                   "MEDIANE BMI",             WHITE),
  list(paste0(nl_num(K$los_med, 1), " d"), "MEDIAAN IZ-VERBLIJF",     LIME)
)
kw <- (SLIDE_W - 2 * M) / length(kpis)
for (i in seq_along(kpis)) {
  x <- M + (i - 1) * kw
  slide1 <- c(slide1, el_kpi(x, 2.34, kw - 0.30, kpis[[i]][[1]], kpis[[i]][[2]], kpis[[i]][[3]], size = 20))
  if (i > 1) slide1 <- c(slide1, list(sp_line(x - 0.15, 2.36, x - 0.15, 2.98, "4A6B7C", 0.75)))
}

## ---- column 1: type of surgery + age distribution --------------------------
slide1 <- c(slide1,
  el_heading(C1, 3.30, "Type ingreep", CW),
  el_hbars(C1, 3.74, CW,
           values = K$surg_dist,
           labels = c("CABG", "CABG + klep", "Klep, geen CABG", "Andere"),
           fills  = c(SLATE, SLATE2, SLATE3, RULE), label_w = 1.52),
  el_heading(C1, 5.50, "Leeftijdsverdeling", CW),
  el_vbars(C1, 5.92, CW, 1.24, K$age_dist, levels(age_band),
           fills = c(SLATE3, SLATE2, SLATE, SLATE2)))

## ---- column 2: comorbidity + narrative -------------------------------------
slide1 <- c(slide1,
  el_heading(C2, 3.30, "Comorbiditeit", CW),
  el_hbars(C2, 3.74, CW,
           values = c(K$htn, K$dm, K$cva, K$lung, K$smoker),
           labels = c("Hypertensie", "Diabetes", "CVA / TIA", "Longlijden", "Actieve roker"),
           fills  = rep(SLATE, 5), label_w = 1.52),
  list(sp_rect(C2, 5.76, CW, 1.40, fill = PANEL, radius = 0.09)),
  list(sp_txt(C2 + 0.22, 5.90, CW - 0.44, 0.26,
              list(para(run("De zorg in het kort", 10.5, INK, bold = TRUE))))),
  list(sp_txt(C2 + 0.22, 6.18, CW - 0.44, 0.90, list(para(run(sprintf(
    paste("Van pre-operatieve risico-evaluatie tot gecontroleerde opwaking op de cardiale IZ.",
          "%s is binnen zes uur ge\u00ebxtubeerd, %s al op tafel; de mediane tijd tot zitten",
          "in de zetel is %s uur."),
    nl_pct(K$extub_6h, 0), nl_pct(K$on_table, 0), nl_num(K$chair_med, 0)), 9.5, MUTED), line = 130)))))

## ---- column 3: baseline table ----------------------------------------------
slide1 <- c(slide1,
  list(sp_rect(C3 - 0.20, 3.18, C3W + 0.20, 3.98, fill = WHITE, line = RULE, lwd = 1, radius = 0.09)),
  el_heading(C3, 3.36, "Basiskarakteristieken", C3W - 0.30),
  el_table(C3, 3.82, C3W - 0.34,
    keys = c("Mediane leeftijd", "Geslacht (man)", "Mediane BMI", "Actieve roker",
             "Diabetes", "Electieve ingreep", "Mediane VAS 24 u", "Tijd tot zetel",
             "Mediaan IZ-verblijf"),
    vals = c(sprintf("%s j.", nl_num(K$age_med, 0)), nl_pct(K$male, 1), nl_num(K$bmi, 1),
             nl_pct(K$smoker, 1), nl_pct(K$dm, 1), nl_pct(K$elective, 1),
             nl_num(K$vas_med, 1), sprintf("%s u", nl_num(K$chair_med, 0)),
             sprintf("%s d", nl_num(K$los_med, 1))),
    accent = c(7, 8, 9)),
  list(sp_txt(C3, 6.74, C3W - 0.34, 0.28, list(para(run(
    sprintf("IQR leeftijd %s j. \u00b7 IQR IZ-verblijf %s d", K$age_iqr, K$los_iqr), 8, MUTED))))))

## ---- footer -----------------------------------------------------------------
slide1 <- c(slide1,
  list(sp_txt(M, 7.20, SLIDE_W - 2 * M - 1.2, 0.24, list(para(run(sprintf(
    "%s IZ-opnames bij %s pati\u00ebnten, %s. Percentages zijn berekend op de geregistreerde waarden.",
    nl_int(K$n_rec), nl_int(K$n_pat), K$period), 7.6, MUTED))))),
  list(sp_txt(SLIDE_W - M - 1.2, 7.20, 1.2, 0.24,
              list(para(run("01", 8, SLATE3, bold = TRUE), align = "r")))))


# =============================================================================
# 6.  SLIDE 2 -- EARLY RECOVERY AND OUTCOMES
# =============================================================================

K$surg_dur <- med(d$surg_dur_min) / 60

slide2 <- c(
  ## ---- header --------------------------------------------------------------
  list(sp_txt(M, 0.34, 8, 0.26, list(para(
    run("CARDIO-ANESTHESIE & IZ", 8.4, MUTED, bold = TRUE, caps = TRUE, spc = 1.8))))),
  list(sp_line(M, 0.74, M + 0.22, 0.74, SLATE, 2.25)),
  list(sp_txt(M + 0.32, 0.58, 9, 0.42,
              list(para(run("Vroeg herstel \u2014 van inductie tot afdeling", 21, INK, bold = TRUE))))),
  list(sp_txt(SLIDE_W - M - 4.6, 0.38, 4.6, 0.24, list(para(
    run("ZORGPAD \u00b7 VROEG HERSTEL \u00b7 UITKOMSTEN", 8.2, MUTED, caps = TRUE, spc = 1.5), align = "r")))),
  list(sp_line(M, 1.14, SLIDE_W - M, 1.14, RULE, 0.75))
)

## ---- care pathway strip ------------------------------------------------------
steps <- list(
  list("Operatiekwartier", "Hemodynamische sturing en peroperatieve monitoring",
       sprintf("%s u ingreep", nl_num(K$surg_dur, 1))),
  list("Extubatie", "Gecontroleerde opwaking, extubatie meestal binnen zes uur",
       sprintf("%s \u2264 6 u", nl_pct(K$extub_6h, 0))),
  list("Cardiale IZ", "Bewaking, pijnstilling en drainverwijdering",
       sprintf("%s d mediaan", nl_num(K$los_med, 1))),
  list("Mobilisatie", "Eerste keer rechtop in de zetel",
       sprintf("%s u tot zetel", nl_num(K$chair_med, 0))),
  list("Ontslag IZ", "Naar medium care of verpleegafdeling",
       sprintf("%s heropname", nl_pct(K$readm, 1)))
)
sw <- (SLIDE_W - 2 * M) / length(steps)
cy <- 1.42; cd <- 0.44
slide2 <- c(slide2, list(sp_line(M + sw / 2, cy + cd / 2, SLIDE_W - M - sw / 2, cy + cd / 2, RULE, 1)))
for (i in seq_along(steps)) {
  cx   <- M + (i - 1) * sw + sw / 2
  last <- i == length(steps)
  slide2 <- c(slide2,
    list(sp_ellipse(cx - cd / 2, cy, cd, cd,
                    fill = if (last) SLATE else WHITE, line = SLATE, lwd = 1.25,
                    paras = list(para(run(as.character(i), 11.5,
                                          if (last) WHITE else SLATE, bold = TRUE), align = "c")))),
    list(sp_txt(cx - sw / 2 + 0.12, cy + cd + 0.08, sw - 0.24, 0.24,
                list(para(run(steps[[i]][[1]], 10.5, INK, bold = TRUE), align = "c")))),
    list(sp_txt(cx - sw / 2 + 0.12, cy + cd + 0.32, sw - 0.24, 0.46,
                list(para(run(steps[[i]][[2]], 8.2, MUTED), align = "c", line = 120)))),
    list(sp_txt(cx - sw / 2 + 0.12, cy + cd + 0.80, sw - 0.24, 0.24,
                list(para(run(steps[[i]][[3]], 9, SLATE, bold = TRUE), align = "c")))))
}

## ---- headline indicator cards -------------------------------------------------
cards <- list(
  list(nl_pct(K$extub_6h, 0), "Ge\u00ebxtubeerd \u2264 6 u", sprintf("%s al op tafel", nl_pct(K$on_table, 0))),
  list(paste0(nl_num(K$los_med, 1), " d"), "Mediaan IZ-verblijf", sprintf("IQR %s d", K$los_iqr)),
  list(nl_pct(K$readm, 1), "Heropname op IZ", "binnen dezelfde opname"),
  list(nl_pct(K$aki_any, 1), "AKI, elk stadium", "creatinine-criterium")
)
cwd <- (SLIDE_W - 2 * M - 3 * 0.24) / 4
for (i in seq_along(cards)) {
  x <- M + (i - 1) * (cwd + 0.24)
  slide2 <- c(slide2,
    list(sp_rect(x, 3.06, cwd, 1.00, fill = WHITE, line = RULE, lwd = 1, radius = 0.09)),
    list(sp_txt(x + 0.20, 3.16, cwd - 0.4, 0.40, list(para(run(cards[[i]][[1]], 22, SLATE, bold = TRUE))))),
    list(sp_txt(x + 0.20, 3.57, cwd - 0.4, 0.22, list(para(run(cards[[i]][[2]], 9.5, INK))))),
    list(sp_txt(x + 0.20, 3.78, cwd - 0.4, 0.20, list(para(run(cards[[i]][[3]], 8, MUTED))))))
}

## ---- left: cumulative extubation curve ---------------------------------------
PX <- M + 0.56; PY <- 4.98; PW <- SLIDE_W * 0.455; PH <- 1.66
slide2 <- c(slide2,
  el_heading(M, 4.34, "Cumulatief aandeel ge\u00ebxtubeerd", 5.4),
  list(sp_txt(M + 0.32, 4.60, 6.4, 0.24,
              list(para(run("tijd sinds einde ingreep (niet-lineaire schaal)", 8.2, MUTED))))))

for (g in c(0, 25, 50, 75, 100)) {
  yy <- PY + PH * (1 - g / 100)
  slide2 <- c(slide2,
    list(sp_line(PX, yy, PX + PW, yy, if (g == 0) RULE else PANEL, if (g == 0) 1 else 0.75)),
    list(sp_txt(PX - 0.56, yy - 0.10, 0.46, 0.20,
                list(para(run(paste0(g, "%"), 7.8, MUTED), align = "r")))))
}
xs <- seq_along(ext_h)
slide2 <- c(slide2, el_step(PX, PY, PW, PH, xs, K$ext_cum, SLATE, 2.25))
for (i in xs) {
  xx <- PX + PW * (i - 1) / (length(xs) - 1)
  slide2 <- c(slide2, list(sp_txt(xx - 0.32, PY + PH + 0.05, 0.64, 0.20,
              list(para(run(paste0(ext_h[i], " u"), 7.8, MUTED), align = "c")))))
}
i6 <- which(ext_h == 6)
x6 <- PX + PW * (i6 - 1) / (length(xs) - 1)
y6 <- PY + PH * (1 - K$ext_cum[i6] / 100)
slide2 <- c(slide2,
  list(sp_line(x6, PY, x6, PY + PH, SLATE3, 1, dash = "dash")),
  list(sp_ellipse(x6 - 0.055, y6 - 0.055, 0.11, 0.11, fill = SLATE)),
  list(sp_rect(x6 + 0.14, y6 + 0.13, 1.46, 0.26, fill = PANEL, radius = 0.06,
               paras = list(para(run(sprintf("%s bij 6 u", nl_pct(K$extub_6h, 0)),
                                     8.5, SLATE, bold = TRUE), align = "c")))))

## ---- right: outcomes ----------------------------------------------------------
RX <- PX + PW + 0.75; RW <- SLIDE_W - M - RX
slide2 <- c(slide2,
  el_heading(RX, 4.34, "Mortaliteit", RW),
  el_hbars(RX, 4.74, RW,
           values = c(K$mort_icu, K$mort_30, K$mort_180),
           labels = c("Op IZ", "30 dagen", "180 dagen"),
           fills  = c(SLATE, SLATE2, SLATE3),
           label_w = 1.30, row_h = 0.30, gap = 0.02,
           max_val = K$mort_180 * 1.2),
  el_heading(RX, 5.86, "AKI-stadium (creatinine)", RW),
  el_hbars(RX, 6.24, RW,
           values = K$aki_dist[2:4],
           labels = c("Stadium 1", "Stadium 2", "Stadium 3"),
           fills  = c(SLATE3, SLATE2, SLATE),
           label_w = 1.30, row_h = 0.30, gap = 0.02,
           max_val = K$aki_dist[2] * 1.2))

## ---- footer -------------------------------------------------------------------
slide2 <- c(slide2,
  list(sp_txt(M, 7.20, SLIDE_W - 2 * M - 1.2, 0.24, list(para(run(sprintf(
    paste("AKI op het creatinine-criterium alleen (%s beoordeelbaar); urineproductie en",
          "nierfunctievervanging ontbreken in de export, de cijfers zijn een ondergrens."),
    nl_int(sum(!is.na(d$aki_stage)))), 7.6, MUTED))))),
  list(sp_txt(SLIDE_W - M - 1.2, 7.20, 1.2, 0.24,
              list(para(run("02", 8, SLATE3, bold = TRUE), align = "r")))))


# =============================================================================
# 7.  BUILD AND EXPORT
# =============================================================================

doc <- officer::add_slide(doc, "Blank", "Office Theme")
doc <- add_shapes(doc, slide1)
doc <- officer::add_slide(doc, "Blank", "Office Theme")
doc <- add_shapes(doc, slide2)

dir.create(dirname(PPTX_OUT), showWarnings = FALSE, recursive = TRUE)
print(doc, target = PPTX_OUT)
cat(sprintf("PowerPoint written: %s  (%.0f kB, 2 slides, %d native shapes)\n",
            PPTX_OUT, file.size(PPTX_OUT) / 1024, length(slide1) + length(slide2)))
