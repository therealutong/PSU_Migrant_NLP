# ============================================================
# Table1_Demographics.R — APA-style demographic Table 1
# ============================================================

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(openxlsx)
  library(flextable)
  library(officer)
})

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

INPUT_PATH  <- "/Users/therealu/Desktop/Ind project/NLP_immigrant/Copy of Respondent Demographic Characteristics for Mental Health Paper.xlsx"
SHEET_NAME  <- "Characteristics"
OUTPUT_DIR  <- "/Users/therealu/Desktop/Ind project/NLP_immigrant"
COMPLETED_ONLY <- TRUE

CATEGORICAL_VARS <- c("Gender", "Visa Status", "Country of Origin",
                      "Educational Attainment", "Employment Status", "Parent?")
CONTINUOUS_VARS  <- c("Age", "Length of Interview (in minutes)")

TABLE_TITLE   <- "Table 1"
TABLE_CAPTION <- "Demographic Characteristics of Respondents"

# Category levels with fewer than this many respondents are collapsed into
# an "Other" bucket (per variable), so a long tail of one-off values (e.g.
# many single-country entries in Country of Origin) doesn't dominate the
# table. Missing and Formula error are never folded into "Other" -- they
# stay as their own labeled rows regardless of size.
MIN_CELL_SIZE <- 5
OTHER_LABEL   <- "Other"

MISSING_LABEL  <- "Missing"
MISSING_TOKENS <- c("", "nan", "none", "n/a", "na", "missing", "unknown", "?")

FORMULA_ERROR_LABEL  <- "Formula error in source file"
FORMULA_ERROR_TOKENS <- c("#value!", "#ref!", "#div/0!", "#name?", "#null!",
                           "#num!", "#n/a", "#error!")

# ---------------------------------------------------------------------------
# Category collapse maps (same collapsing as the Python version)
# Names are normalized (lowercase, whitespace-collapsed) raw values;
# values are the clean label they should be grouped under.
# ---------------------------------------------------------------------------

VISA_STATUS_MAP <- c(
  "f-1" = "F-1", "f-2" = "F-2", "h-1b" = "H-1B", "h-4" = "H-4",
  "green card" = "Green Card", "j-1" = "J-1", "j-2" = "J-2",
  "f-1 (opt)" = "F-1",
  "f-1, currently working so out of status" = "F-1 (out of status)",
  "h-1b (previously f-1 and opt)" = "H-1B",
  "j-1 (previously f-1 and f-1 opt)" = "J-1",
  "5-year b-1/b-2" = "B-1/B-2"
)

EDUCATION_MAP <- c(
  "bachelor's" = "Bachelor's", "bachelors" = "Bachelor's",
  "bachelor's, lato sensu post graduate studies" = "Bachelor's",
  "bachelor's, master's (in progress)" = "Master's",
  "master's" = "Master's", "master's (in progress)" = "Master's",
  "master's (starting soon)" = "Master's", "master's (online)" = "Master's",
  "phd" = "Doctoral/MD", "phd (in progress)" = "Doctoral/MD",
  "phd (not in u.s., online)" = "Doctoral/MD",
  "md in ghana" = "Doctoral/MD", "equivalent of md in ghana" = "Doctoral/MD",
  "md?" = "Doctoral/MD", "m.d" = "Doctoral/MD",
  "no higher education?" = "No higher education",
  "higher secondry?" = "Secondary/high school"
)

EMPLOYMENT_MAP <- c(
  "employed" = "Employed", "unemployed" = "Unemployed",
  "work-study" = "Work-Study", "self-employed" = "Self-Employed",
  "student" = "Student",
  "employed (under the table?)" = "Employed",
  "employed (under the table)" = "Employed"
)

PARENT_MAP <- c(
  "no" = "No", "yes - 1" = "Yes", "yes - 2" = "Yes",
  "yes - 3" = "Yes", "yes - 4" = "Yes",
  "pregnant (1st)" = "Pregnant (expecting first child)"
)

COUNTRY_MAP <- c("brazil/ peru" = "Brazil/Peru")

CATEGORY_MAPS <- list(
  "Visa Status" = VISA_STATUS_MAP,
  "Educational Attainment" = EDUCATION_MAP,
  "Employment Status" = EMPLOYMENT_MAP,
  "Parent?" = PARENT_MAP,
  "Country of Origin" = COUNTRY_MAP
)

# Variables where the map above is meant to cover every real value; anything
# falling through here is worth a console warning. Country of Origin and any
# variable without a map (e.g. Gender) have many valid pass-through values
# and are not checked.
FULL_COVERAGE_VARS <- c("Visa Status", "Educational Attainment",
                         "Employment Status", "Parent?")

normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- str_squish(x)
  tolower(x)
}

`%||%` <- function(a, b) if (is.null(a)) b else a

# ---------------------------------------------------------------------------
# Step 1: read the sheet, filling merged cells
# ---------------------------------------------------------------------------
# Row 1 is a title row; row 2 holds the real headers, so startRow = 2 makes
# row 2 the column names and data begin at row 3. fillMergedCells = TRUE
# propagates a merged range's value (e.g. a couple sharing one Country of
# Origin cell) into every row it covers, instead of leaving every row but
# the first blank.
df <- openxlsx::read.xlsx(INPUT_PATH, sheet = SHEET_NAME, startRow = 2,
                           fillMergedCells = TRUE, detectDates = FALSE,
                           check.names = FALSE)
names(df) <- trimws(names(df))
df <- df[rowSums(!is.na(df) & df != "") > 0, , drop = FALSE]

# ---------------------------------------------------------------------------
# Column-name matching, robust to whitespace variants
# ---------------------------------------------------------------------------
# Excel headers sometimes carry a non-breaking space, a double space, or
# trailing whitespace that plain trimws() won't catch (trimws only strips
# ASCII space/tab/newline, not U+00A0). An exact string match against
# CATEGORICAL_VARS/CONTINUOUS_VARS can then silently fail -- the variable
# just never appears in the table, with no error. Column lookup below
# normalizes both sides (collapse all whitespace, including U+00A0, to a
# single space; lowercase) before matching, and prints the raw column
# names it found so any real mismatch is visible immediately.
normalize_colname <- function(x) {
  x <- gsub("\u00A0", " ", x)   # non-breaking space -> regular space
  # R's default data.frame name-mangling turns spaces into dots (e.g.
  # "Visa Status" -> "Visa.Status") if check.names ever ends up TRUE
  # (different openxlsx versions/environments can behave differently
  # here even when we ask for check.names = FALSE). Treat a dot the same
  # as a space so matching survives either way.
  x <- gsub("\\.", " ", x)
  x <- str_squish(x)
  tolower(x)
}

cat("Columns found in the sheet:\n")
for (nm in names(df)) cat(sprintf("  %s\n", nm))

col_lookup <- setNames(names(df), normalize_colname(names(df)))

resolve_column <- function(target) {
  key <- normalize_colname(target)
  if (key %in% names(col_lookup)) return(unname(col_lookup[[key]]))
  NA_character_
}

# Resolve each configured variable name to its actual column name once, up
# front, and warn about any that couldn't be matched at all (as opposed to
# matching but being empty/missing data, which is a separate, normal case).
resolve_and_warn <- function(vars, kind) {
  resolved <- vapply(vars, resolve_column, character(1))
  unresolved <- vars[is.na(resolved)]
  if (length(unresolved)) {
    cat(sprintf("\n[warning] %s variable(s) not found in the sheet (check spelling/whitespace):\n", kind))
    for (v in unresolved) cat(sprintf("  - %s\n", v))
  }
  setNames(resolved, vars)
}

CATEGORICAL_COLS <- resolve_and_warn(CATEGORICAL_VARS, "categorical")
CONTINUOUS_COLS  <- resolve_and_warn(CONTINUOUS_VARS, "continuous")

# ---------------------------------------------------------------------------
# Step 2: recover formula-error cells where possible (optional; needs tidyxl)
# ---------------------------------------------------------------------------
# A cell built from a formula (e.g. one rendering a flag icon next to the
# country name, via a Geography data type or an IMAGE() lookup) can cache
# an error string like "#VALUE!" if that lookup fails. Where the real text
# is still recoverable from the formula, patch it back in; otherwise leave
# it for the FORMULA_ERROR_TOKENS handling below and report it.

try_recover_from_formula <- function(formula_text, resolve_ref) {
  if (is.na(formula_text) || !nzchar(formula_text)) return(NA_character_)

  # Heuristic 1: a quoted literal that isn't a URL
  candidates <- unlist(str_extract_all(formula_text, '"([^"]+)"'))
  candidates <- gsub('^"|"$', "", candidates)
  candidates <- trimws(candidates)
  candidates <- candidates[nzchar(candidates)]
  candidates <- candidates[!grepl("http|www\\.", candidates, ignore.case = TRUE)]
  candidates <- candidates[grepl("[A-Za-z]{3,}", candidates)]
  if (length(candidates)) return(candidates[which.max(nchar(candidates))])

  # Heuristic 2: formula is a single cell reference, e.g. D5 or $D$5
  m <- regmatches(formula_text,
                   regexec("^\\$?([A-Za-z]{1,3})\\$?([0-9]{1,7})$", formula_text))[[1]]
  if (length(m) == 3) return(resolve_ref(m[2], as.integer(m[3])))

  NA_character_
}

recovered_log  <- list()
unresolved_log <- list()

if (requireNamespace("tidyxl", quietly = TRUE)) {
  cells <- tidyxl::xlsx_cells(INPUT_PATH, sheets = SHEET_NAME)

  cell_value <- function(row_i) {
    r <- cells[row_i, ]
    switch(r$data_type,
           "character" = r$character,
           "numeric"   = as.character(r$numeric),
           "date"      = as.character(r$date),
           "logical"   = as.character(r$logical),
           "error"     = r$error,
           NA_character_)
  }

  by_addr <- split(seq_len(nrow(cells)), cells$address)

  resolve_ref <- function(col_letters, row_number, visited = character(0)) {
    addr <- paste0(toupper(col_letters), row_number)
    if (addr %in% visited || is.null(by_addr[[addr]])) return(NA_character_)
    visited <- c(visited, addr)
    i <- by_addr[[addr]][1]
    val <- cell_value(i)
    if (!is.na(val) && normalize_key(val) %in% FORMULA_ERROR_TOKENS) {
      ref_formula <- cells$formula[i]
      return(try_recover_from_formula(
        ref_formula, function(cl, rn) resolve_ref(cl, rn, visited)))
    }
    val
  }

  error_rows <- which(cells$data_type == "error")
  for (i in error_rows) {
    formula_text <- cells$formula[i]
    recovered <- try_recover_from_formula(formula_text, resolve_ref)
    addr <- cells$address[i]
    if (!is.na(recovered)) {
      recovered_log[[length(recovered_log) + 1]] <- list(addr = addr, old = cells$error[i], new = recovered)
      # patch df: map address -> data-frame row/col (data starts at sheet row 3)
      col_letters <- gsub("[0-9]", "", addr)
      row_number  <- as.integer(gsub("[A-Za-z]", "", addr))
      col_index   <- openxlsx::col2int(col_letters)
      df_row <- row_number - 2   # header is sheet row 2, data starts row 3
      if (df_row >= 1 && df_row <= nrow(df) && col_index <= ncol(df)) {
        df[df_row, col_index] <- recovered
      }
    } else {
      unresolved_log[[length(unresolved_log) + 1]] <- list(addr = addr, formula = formula_text)
    }
  }
} else {
  message("[note] package 'tidyxl' not installed -- skipping automatic formula-",
          "error recovery. Cells with cached errors (e.g. #VALUE!) will be ",
          "labeled '", FORMULA_ERROR_LABEL, "' without attempting recovery. ",
          "Install tidyxl to enable recovery: install.packages('tidyxl')")
}

if (length(recovered_log)) {
  cat("\nRecovered the real value from a broken formula:\n")
  for (r in recovered_log) cat(sprintf("  %s: %s -> %s\n", r$addr, r$old, r$new))
}
if (length(unresolved_log)) {
  cat(sprintf('\nCould not automatically recover these formula-error cells (shown as "%s"):\n',
              FORMULA_ERROR_LABEL))
  for (r in unresolved_log) cat(sprintf("  %s: %s\n", r$addr, r$formula))
}

# ---------------------------------------------------------------------------
# Step 3: filter to completed interviews, if requested
# ---------------------------------------------------------------------------
completed_col <- resolve_column("Completed?")
if (COMPLETED_ONLY && !is.na(completed_col)) {
  flag <- tolower(trimws(as.character(df[[completed_col]])))
  df <- df[flag %in% c("yes", "y", "complete", "completed", "1", "true"), , drop = FALSE]
}
N_TOTAL <- nrow(df)

# ---------------------------------------------------------------------------
# Step 4: coerce continuous variables to numeric
# ---------------------------------------------------------------------------
for (var in CONTINUOUS_VARS) {
  col <- CONTINUOUS_COLS[[var]]
  if (!is.na(col)) df[[col]] <- suppressWarnings(as.numeric(df[[col]]))
}

# ---------------------------------------------------------------------------
# Step 5: collapse categorical values
# ---------------------------------------------------------------------------
clean_categorical <- function(x, map) {
  key <- normalize_key(x)
  out <- character(length(x))
  unmapped <- character(0)
  n_formula_errors <- 0

  for (i in seq_along(x)) {
    k <- key[i]
    if (k %in% FORMULA_ERROR_TOKENS) {
      out[i] <- FORMULA_ERROR_LABEL
      n_formula_errors <- n_formula_errors + 1
    } else if (k %in% MISSING_TOKENS || is.na(x[i])) {
      out[i] <- MISSING_LABEL
    } else if (k %in% names(map)) {
      out[i] <- unname(map[[k]])
    } else {
      passthrough <- trimws(as.character(x[i]))
      out[i] <- passthrough
      unmapped <- c(unmapped, passthrough)
    }
  }
  list(cleaned = out, unmapped = unique(unmapped), n_formula_errors = n_formula_errors)
}

# ---------------------------------------------------------------------------
# Step 6: build the table rows
# ---------------------------------------------------------------------------
rows <- list()
all_unmapped <- list()
all_formula_errors <- list()

add_row <- function(variable, unit, value, range = "", is_section = FALSE) {
  rows[[length(rows) + 1]] <<- data.frame(
    Variable = variable, Unit = unit, Value = value, Range = range,
    is_section = is_section, stringsAsFactors = FALSE
  )
}

for (var in CATEGORICAL_VARS) {
  col <- CATEGORICAL_COLS[[var]]
  if (is.na(col)) next
  res <- clean_categorical(df[[col]], CATEGORY_MAPS[[var]] %||% list())
  if (length(res$unmapped) && var %in% FULL_COVERAGE_VARS) all_unmapped[[var]] <- res$unmapped
  if (res$n_formula_errors) all_formula_errors[[var]] <- res$n_formula_errors

  add_row(var, "", "", is_section = TRUE)
  tb <- sort(table(res$cleaned), decreasing = TRUE)

  # Collapse small cells into "Other" (never Missing / Formula error).
  protected <- c(MISSING_LABEL, FORMULA_ERROR_LABEL)
  small <- names(tb)[tb < MIN_CELL_SIZE & !(names(tb) %in% protected)]
  if (length(small)) {
    other_n <- sum(tb[small])
    tb <- tb[!(names(tb) %in% small)]
    if (OTHER_LABEL %in% names(tb)) {
      tb[[OTHER_LABEL]] <- tb[[OTHER_LABEL]] + other_n
    } else {
      tb <- c(tb, setNames(other_n, OTHER_LABEL))
    }
    tb <- sort(tb, decreasing = TRUE)
  }

  tail_labels <- intersect(c(OTHER_LABEL, MISSING_LABEL, FORMULA_ERROR_LABEL), names(tb))
  ordered <- c(setdiff(names(tb), tail_labels), tail_labels)
  for (lv in ordered) {
    n <- tb[[lv]]
    pct <- if (N_TOTAL > 0) 100 * n / N_TOTAL else 0
    add_row(paste0("  ", lv), "%", sprintf("%d", n), sprintf("%.1f%%", pct))
  }
}

for (var in CONTINUOUS_VARS) {
  col <- CONTINUOUS_COLS[[var]]
  if (is.na(col)) next
  x <- df[[col]]; x <- x[is.finite(x)]
  if (!length(x)) {
    cat(sprintf("\n[warning] '%s' matched column '%s' but had no usable numeric values ",
                var, col))
    cat("(check for text, blanks, or a wrong column match) -- omitted from the table.\n")
    next
  }
  add_row(var, "", "", is_section = TRUE)
  add_row("  M (SD)", "M (SD)", sprintf("%.1f (%.1f)", mean(x), sd(x)),
          sprintf("Range: %.0f\u2013%.0f", min(x), max(x)))
}

T1 <- bind_rows(rows)
sec_at <- which(T1$is_section)
T1$is_section <- NULL

# ---------------------------------------------------------------------------
# Step 7: console warnings
# ---------------------------------------------------------------------------
if (length(all_formula_errors)) {
  cat("\nExcel formula errors found (e.g. #VALUE!), counted separately from\n")
  cat("real missing data:\n")
  for (v in names(all_formula_errors)) cat(sprintf("  %s: %d cell(s)\n", v, all_formula_errors[[v]]))
  cat("\nRecommended fix: in the source workbook, select the affected column,\n")
  cat("copy it, then Paste Special > Values Only over itself.\n")
}
if (length(all_unmapped)) {
  cat("\nValues not found in a collapse map (passed through as-is):\n")
  for (v in names(all_unmapped)) {
    cat(sprintf("  %s:\n", v))
    for (val in sort(all_unmapped[[v]])) cat(sprintf("    - %s\n", val))
  }
  cat("\nIf any of these are variants of an existing category, add them to the\n")
  cat("relevant *_MAP vector above and re-run.\n")
}

cat(sprintf("\nN included: %d\n", N_TOTAL))
print(T1, row.names = FALSE)

# ---------------------------------------------------------------------------
# Step 8: write the .docx (flextable + officer, mirroring DD_Wave1_Descriptives.R)
# ---------------------------------------------------------------------------
note <- sprintf("Note. N = %d.", N_TOTAL)

ft <- flextable(T1) %>%
  theme_booktabs() %>%
  hline(i = 1, j = 2:4, part = "header", border = fp_border(color = "black", width = 1)) %>%
  bold(i = sec_at, j = 1, part = "body") %>%
  padding(i = setdiff(seq_len(nrow(T1)), sec_at), j = 1, padding.left = 12, part = "body") %>%
  align(j = 2:4, align = "center", part = "all") %>%
  align(j = 1, align = "left", part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  fontsize(size = 11, part = "all") %>%
  autofit() %>%
  add_footer_lines(note) %>%
  fontsize(size = 9, part = "footer")

docx_path <- file.path(OUTPUT_DIR, "Table1_Demographics.docx")
print(read_docx() %>% body_add_par(TABLE_TITLE, style = "heading 1") %>%
        body_add_par(TABLE_CAPTION, style = "Normal") %>%
        body_add_flextable(ft),
      target = docx_path)
cat(sprintf("\nWrote %s\n", docx_path))

# ---------------------------------------------------------------------------
# Step 9: write the .tex (self-contained APA three-line table, booktabs style)
# ---------------------------------------------------------------------------
tex_escape <- function(x) {
  x <- gsub("\\", "\u0001BS\u0001", x, fixed = TRUE)  # placeholder, replaced back last
  x <- gsub("%", "\\%", x, fixed = TRUE)
  x <- gsub("$", "\\$", x, fixed = TRUE)
  x <- gsub("#", "\\#", x, fixed = TRUE)
  x <- gsub("&", "\\&", x, fixed = TRUE)
  x <- gsub("_", "\\_", x, fixed = TRUE)
  x <- gsub("{", "\\{", x, fixed = TRUE)
  x <- gsub("}", "\\}", x, fixed = TRUE)
  x <- gsub("\u0001BS\u0001", "\\textbackslash{}", x, fixed = TRUE)
  x
}

write_apa_tex <- function(T1, sec_at, caption, note, path) {
  lines <- c(
    "\\begin{table}[!htb]",
    "\\centering",
    sprintf("\\caption{%s}", tex_escape(caption)),
    "\\begin{tabular}{lccc}",
    "\\toprule",
    "Characteristic & Unit & Value & Range \\\\",
    "\\midrule"
  )
  for (i in seq_len(nrow(T1))) {
    var <- tex_escape(T1$Variable[i])
    if (i %in% sec_at) {
      lines <- c(lines, sprintf("\\textbf{%s} & & & \\\\", var))
    } else {
      lines <- c(lines, sprintf("\\hspace{1em}%s & %s & %s & %s \\\\",
                                 var, tex_escape(as.character(T1$Unit[i])),
                                 tex_escape(as.character(T1$Value[i])),
                                 tex_escape(as.character(T1$Range[i]))))
    }
  }
  lines <- c(lines, "\\bottomrule",
             sprintf("\\multicolumn{4}{l}{\\footnotesize %s} \\\\", tex_escape(note)),
             "\\end{tabular}",
             "\\end{table}")
  writeLines(lines, path)
}

tex_path <- file.path(OUTPUT_DIR, "Table1_Demographics.tex")
write_apa_tex(T1, sec_at, TABLE_CAPTION, note, tex_path)
cat(sprintf("Wrote %s\n", tex_path))
