library(data.table)
library(readxl)
library(jsonlite)

project_dir <- normalizePath(file.path(Sys.getenv("HOME"), "dataset-codebooks"))

requested_file <- "/Users/sungjunpark/Downloads/표본코호트DB 2.2 레이아웃.xlsx"
source_file <- file.path(project_dir, "assets", "맞춤형 자료 제공 컬럼 레이아웃_2026_v1.xlsx")
source_sheet <- "맞춤형 제공 컬럼"
supplement_sheet <- "사업장업종세분류"

out_files <- file.path(project_dir, "index.html")

clean_text <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("\r\n|\r", "\n", x)
  x <- gsub("[ \t]+", " ", x)
  x <- gsub(" *\n *", "\n", x)
  trimws(x)
}

clean_table <- function(x) {
  x <- clean_text(x)
  x <- gsub("\n", " ", x)
  gsub(" +", " ", x)
}

is_blank <- function(x) {
  is.na(x) | trimws(as.character(x)) == ""
}

unique_nonblank <- function(x) {
  x <- clean_text(x)
  unique(x[x != ""])
}

fill_down_character <- function(x) {
  out <- as.character(x)
  last <- NA_character_
  for (i in seq_along(out)) {
    if (!is.na(out[i]) && out[i] != "") {
      last <- out[i]
    }
    out[i] <- last
  }
  out
}

read_clean_sheet <- function(path, sheet) {
  DT <- as.data.table(suppressMessages(
    read_excel(path, sheet = sheet, col_names = FALSE, col_types = "text")
  ))
  setnames(DT, paste0("c", seq_len(ncol(DT))))
  for (nm in names(DT)) {
    DT[, (nm) := clean_text(.SD[[1]]), .SDcols = nm]
  }
  DT[, rid := .I]
  DT
}

extract_value_pairs <- function(block) {
  value_cols <- paste0("c", c(6, 8, 10, 12, 14))
  label_cols <- paste0("c", c(7, 9, 11, 13, 15))

  values <- rbindlist(lapply(seq_along(value_cols), function(i) {
    data.table(
      rid = block$rid,
      pair_order = i,
      value = block[[value_cols[i]]],
      label = block[[label_cols[i]]]
    )
  }))

  values <- values[!(is_blank(value) & is_blank(label))]
  if (nrow(values) == 0) {
    return(list())
  }

  setorder(values, rid, pair_order)
  values <- values[, .(value = clean_text(value), label = clean_text(label))]
  values <- unique(values, by = c("value", "label"))
  lapply(seq_len(nrow(values)), function(i) {
    list(value = values$value[i], label = values$label[i])
  })
}

extract_business_codes <- function(path, sheet) {
  DT <- read_clean_sheet(path, sheet)
  pair_list <- list(
    DT[6:.N, .(value = c1, label = c2)],
    DT[6:.N, .(value = c3, label = c4)],
    DT[6:.N, .(value = c5, label = c6)]
  )
  values <- rbindlist(pair_list)
  values <- values[!(is_blank(value) & is_blank(label))]
  values <- values[, .(value = clean_text(value), label = clean_text(label))]
  values <- unique(values, by = c("value", "label"))
  lapply(seq_len(nrow(values)), function(i) {
    list(value = values$value[i], label = values$label[i])
  })
}

main <- read_clean_sheet(source_file, source_sheet)

main[, table_marker := fifelse(
  !is_blank(c1) & !c1 %in% c("맞춤형 제공 테이블 레이아웃", "테이블 구분"),
  clean_table(c1),
  NA_character_
)]
main[, table_group := fill_down_character(table_marker)]

variable_starts <- main[
  rid > 3 &
    !is_blank(c3) &
    !is_blank(c4) &
    c3 != "순번" &
    c4 != "변수 명",
  rid
]

business_codes <- extract_business_codes(source_file, supplement_sheet)

variables <- lapply(seq_along(variable_starts), function(i) {
  start <- variable_starts[i]
  end <- if (i < length(variable_starts)) variable_starts[i + 1] - 1 else nrow(main)
  block <- main[rid >= start & rid <= end]
  first <- block[1]
  notes <- unique_nonblank(block$c16)

  list(
    table = first$table_group,
    seq = first$c3,
    variable = first$c4,
    description = first$c5,
    values = extract_value_pairs(block),
    business_values = if (first$c4 == "INDTP_CD") business_codes else list(),
    notes = as.list(notes)
  )
})

requested_note <- ""
if (file.exists(requested_file)) {
  requested_sheets <- excel_sheets(requested_file)
  if (!source_sheet %in% requested_sheets) {
    requested_note <- paste0(
      "요청 파일 '", basename(requested_file), "'에는 '", source_sheet,
      "' 시트가 없어 동일 폴더의 '", basename(source_file), "'을 사용했습니다."
    )
  }
}

payload <- list(
  meta = list(
    title = "표본코호트 맞춤형 제공 컬럼",
    source_file = basename(source_file),
    source_sheet = source_sheet,
    supplement_sheet = supplement_sheet,
    requested_file = basename(requested_file),
    requested_note = requested_note,
    generated_at = format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z"),
    total_variables = length(variables)
  ),
  variables = variables
)

payload_json <- toJSON(payload, auto_unbox = TRUE, pretty = FALSE, na = "null")
payload_json <- gsub("</", "<\\/", payload_json, fixed = TRUE)

html_escape <- function(x, keep_breaks = TRUE) {
  x <- clean_text(x)
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x <- gsub('"', "&quot;", x, fixed = TRUE)
  x <- gsub("'", "&#039;", x, fixed = TRUE)
  if (keep_breaks) {
    x <- gsub("\n", "<br>", x, fixed = TRUE)
  } else {
    x <- gsub("\n", " ", x, fixed = TRUE)
  }
  x
}

render_value_table_static <- function(values) {
  if (length(values) == 0) {
    return('<div class="empty">-</div>')
  }

  rows <- vapply(values, function(item) {
    value <- if (item$value == "") "-" else item$value
    label <- if (item$label == "") "-" else item$label
    paste0(
      '<tr>',
      '<td class="value-code" data-label="변수값">', html_escape(value), '</td>',
      '<td class="value-label" data-label="변수값설명">', html_escape(label), '</td>',
      '</tr>'
    )
  }, character(1))

  paste0(
    '<div class="values-wrap"><table><thead><tr>',
    '<th>변수값</th><th>변수값설명</th>',
    '</tr></thead><tbody>',
    paste(rows, collapse = ""),
    '</tbody></table></div>'
  )
}

render_card_static <- function(item) {
  note_html <- ""
  notes <- unlist(item$notes, use.names = FALSE)
  if (length(notes) > 0) {
    note_html <- paste0(
      '<div class="note"><strong>비고</strong><br>',
      paste(html_escape(notes), collapse = "<br>"),
      '</div>'
    )
  }

  business_html <- ""
  if (length(item$business_values) > 0) {
    business_html <- paste0(
      '<section class="extra-values">',
      '<div class="section-title">사업장업종세분류 상세 코드 <b>',
      format(length(item$business_values), big.mark = ","),
      '개</b></div>',
      render_value_table_static(item$business_values),
      '</section>'
    )
  }

  paste0(
    '<article class="card" data-table="', html_escape(item$table, keep_breaks = FALSE),
    '" data-variable="', html_escape(item$variable, keep_breaks = FALSE), '">',
    '<div class="card-head"><div>',
    '<div class="kicker">', html_escape(item$table), '</div>',
    '<div class="var-line">',
    '<div class="var-name">', html_escape(item$variable), '</div>',
    '<div class="seq">순번 ', html_escape(item$seq), '</div>',
    '</div>',
    '<div class="desc">', html_escape(item$description), '</div>',
    '</div>', note_html, '</div>',
    '<div class="card-body"><section>',
    '<div class="section-title">변수 세부설명 <b>',
    format(length(item$values), big.mark = ","),
    '개</b></div>',
    render_value_table_static(item$values),
    '</section>',
    business_html,
    '</div></article>'
  )
}

source_note_html <- paste0(
  '<strong>참고 파일</strong> ', html_escape(payload$meta$source_file),
  '<span> / ', html_escape(payload$meta$source_sheet), ', ',
  html_escape(payload$meta$supplement_sheet), '</span>',
  if (payload$meta$requested_note != "") {
    paste0('<p>', html_escape(payload$meta$requested_note), '</p>')
  } else {
    ""
  }
)

cards_html <- paste(vapply(variables, render_card_static, character(1)), collapse = "\n")

html_head <- paste0('<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="robots" content="noindex, nofollow">
  <title>표본코호트 맞춤형 제공 컬럼</title>
  <style>
    @font-face {
      font-family: "PretendardVariable";
      src: url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard/packages/pretendard/dist/web/variable/woff2/PretendardVariable.woff2") format("woff2");
      font-weight: 100 900;
      font-display: swap;
    }

    :root {
      color-scheme: light;
      --bg: #f8f9fa;
      --panel: #ffffff;
      --panel-subtle: #fbfcfd;
      --ink: #212529;
      --muted: #5f6874;
      --line: #dee2e6;
      --soft: #e9ecef;
      --accent: #d0473e;
      --accent-strong: #a9352d;
      --accent-soft: #fff0ed;
      --focus: rgba(208, 71, 62, .24);
      --shadow: 0 1px 2px rgba(33, 37, 41, .05), 0 12px 34px rgba(33, 37, 41, .07);
    }

    :root[data-theme="dark"] {
      color-scheme: dark;
      --bg: #212529;
      --panel: #2b3035;
      --panel-subtle: #343a40;
      --ink: #ebebeb;
      --muted: #b9c0c8;
      --line: #4e5d6c;
      --soft: #343a40;
      --accent: #df6919;
      --accent-strong: #f08a3c;
      --accent-soft: rgba(223, 105, 25, .16);
      --focus: rgba(223, 105, 25, .35);
      --shadow: 0 1px 2px rgba(0, 0, 0, .16), 0 14px 38px rgba(0, 0, 0, .18);
    }

    @media (prefers-color-scheme: dark) {
      :root:not([data-theme="light"]) {
        color-scheme: dark;
        --bg: #212529;
        --panel: #2b3035;
        --panel-subtle: #343a40;
        --ink: #ebebeb;
        --muted: #b9c0c8;
        --line: #4e5d6c;
        --soft: #343a40;
        --accent: #df6919;
        --accent-strong: #f08a3c;
        --accent-soft: rgba(223, 105, 25, .16);
        --focus: rgba(223, 105, 25, .35);
        --shadow: 0 1px 2px rgba(0, 0, 0, .16), 0 14px 38px rgba(0, 0, 0, .18);
      }
    }

    * {
      box-sizing: border-box;
    }

    html {
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }

    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: "PretendardVariable", -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Noto Sans KR", "Segoe UI", sans-serif;
      font-size: 16px;
      line-height: 1.55;
    }

    button,
    select {
      font: inherit;
    }

    .page {
      width: min(1120px, calc(100% - 28px));
      margin: 0 auto;
    }

    header {
      position: sticky;
      top: 0;
      z-index: 20;
      border-bottom: 1px solid var(--line);
      background: color-mix(in srgb, var(--bg) 94%, transparent);
      backdrop-filter: blur(10px);
    }

    .top {
      padding: 18px 0 14px;
    }

    .title-row {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 14px;
    }

    h1 {
      margin: 0;
      font-size: clamp(1.35rem, 2.7vw, 2rem);
      line-height: 1.2;
      font-weight: 800;
      letter-spacing: 0;
      word-break: keep-all;
    }

    .count-pill {
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 5px 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      color: var(--muted);
      font-size: .86rem;
      font-weight: 700;
      white-space: nowrap;
    }

    .count-pill b {
      color: var(--accent);
      font-weight: 850;
    }

    .controls {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) 150px;
      gap: 10px;
    }

    .control {
      min-width: 0;
    }

    .control label {
      display: block;
      margin-bottom: 5px;
      color: var(--muted);
      font-size: .78rem;
      font-weight: 800;
    }

    select {
      width: 100%;
      min-height: 44px;
      padding: 8px 38px 8px 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      color: var(--ink);
      outline: none;
    }

    select:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 4px var(--focus);
    }

    main {
      padding: 18px 0 26px;
    }

    .source-note {
      margin: 0 0 12px;
      padding: 12px 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      color: var(--muted);
      font-size: .88rem;
    }

    .source-note strong {
      color: var(--ink);
    }

    .source-note p {
      margin: 4px 0 0;
    }

    .cards {
      display: grid;
      gap: 12px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .card-head {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 12px;
      padding: 15px 16px 13px;
      border-bottom: 1px solid var(--line);
      background: var(--panel-subtle);
    }

    .kicker {
      margin-bottom: 4px;
      color: var(--accent);
      font-size: .78rem;
      font-weight: 850;
      word-break: keep-all;
    }

    .var-line {
      display: flex;
      flex-wrap: wrap;
      align-items: baseline;
      gap: 8px 10px;
    }

    .var-name {
      font-family: "SFMono-Regular", "Menlo", "Consolas", monospace;
      font-size: clamp(1.08rem, 2.2vw, 1.35rem);
      font-weight: 850;
      letter-spacing: 0;
      word-break: break-word;
    }

    .seq {
      color: var(--muted);
      font-size: .82rem;
      font-weight: 800;
      white-space: nowrap;
    }

    .desc {
      margin-top: 7px;
      font-size: 1rem;
      font-weight: 750;
      word-break: keep-all;
    }

    .note {
      align-self: start;
      max-width: 280px;
      padding: 7px 9px;
      border: 1px solid color-mix(in srgb, var(--accent) 45%, var(--line));
      border-radius: 8px;
      background: var(--accent-soft);
      color: var(--ink);
      font-size: .82rem;
      font-weight: 700;
      word-break: keep-all;
    }

    .card-body {
      padding: 14px 16px 16px;
    }

    .section-title {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      margin: 0 0 8px;
      color: var(--muted);
      font-size: .82rem;
      font-weight: 850;
    }

    .section-title b {
      color: var(--accent);
      font-weight: 850;
    }

    .values-wrap {
      max-height: 19rem;
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
    }

    .card.is-focused .values-wrap {
      max-height: 42rem;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: .9rem;
    }

    th,
    td {
      padding: 8px 10px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: var(--soft);
      color: var(--ink);
      font-size: .78rem;
      font-weight: 850;
    }

    tr:last-child td {
      border-bottom: 0;
    }

    .value-code {
      width: 28%;
      min-width: 130px;
      font-family: "SFMono-Regular", "Menlo", "Consolas", monospace;
      font-weight: 780;
      word-break: break-word;
    }

    .value-label {
      word-break: keep-all;
      overflow-wrap: anywhere;
    }

    .empty {
      padding: 10px;
      color: var(--muted);
      font-size: .9rem;
    }

    .extra-values {
      margin-top: 13px;
    }

    .no-results {
      display: none;
      padding: 18px;
      border: 1px dashed var(--line);
      border-radius: 8px;
      color: var(--muted);
      text-align: center;
    }

    .no-results.is-visible {
      display: block;
    }

    @media (max-width: 760px) {
      .page {
        width: min(100% - 20px, 1120px);
      }

      .top {
        padding: 13px 0 11px;
      }

      .title-row {
        display: block;
        margin-bottom: 10px;
      }

      .count-pill {
        margin-top: 8px;
      }

      .controls {
        grid-template-columns: 1fr;
        gap: 8px;
      }

      main {
        padding-top: 12px;
      }

      .card-head {
        grid-template-columns: 1fr;
        padding: 13px;
      }

      .note {
        max-width: none;
      }

      .card-body {
        padding: 12px 13px 13px;
      }

      table,
      tbody,
      tr,
      td {
        display: block;
        width: 100%;
      }

      thead {
        display: none;
      }

      tr {
        padding: 8px 10px;
        border-bottom: 1px solid var(--line);
      }

      tr:last-child {
        border-bottom: 0;
      }

      td {
        padding: 3px 0;
        border-bottom: 0;
      }

      td::before {
        display: block;
        margin-bottom: 1px;
        color: var(--muted);
        font-size: .74rem;
        font-weight: 850;
        content: attr(data-label);
      }

      .value-code {
        min-width: 0;
      }
    }
  </style>
</head>
<body>
  <header>
    <div class="page top">
      <div class="title-row">
        <h1>표본코호트 맞춤형 제공 컬럼</h1>
        <div class="count-pill">표시 변수 <b id="visibleCount">', length(variables), '</b> / <span id="totalCount">', length(variables), '</span></div>
      </div>
      <div class="controls">
        <div class="control">
          <label for="tableFilter">테이블 구분</label>
          <select id="tableFilter"></select>
        </div>
        <div class="control">
          <label for="variableFilter">변수명</label>
          <select id="variableFilter"></select>
        </div>
        <div class="control">
          <label for="themeSelect">화면 모드</label>
          <select id="themeSelect">
            <option value="auto">자동</option>
            <option value="light">라이트</option>
            <option value="dark">다크</option>
          </select>
        </div>
      </div>
    </div>
  </header>
  <main class="page">
    <div class="source-note" id="sourceNote">', source_note_html, '</div>
    <section class="cards" id="cards">', cards_html, '</section>
    <div class="no-results" id="noResults">선택한 조건에 맞는 변수가 없습니다.</div>
  </main>
  <script type="application/json" id="payload">')

html_tail <- '</script>
  <script>
    const payload = JSON.parse(document.getElementById("payload").textContent);
    const variables = payload.variables;
    const meta = payload.meta;

    const tableFilter = document.getElementById("tableFilter");
    const variableFilter = document.getElementById("variableFilter");
    const themeSelect = document.getElementById("themeSelect");
    const cards = document.getElementById("cards");
    const noResults = document.getElementById("noResults");
    const visibleCount = document.getElementById("visibleCount");
    const totalCount = document.getElementById("totalCount");
    const sourceNote = document.getElementById("sourceNote");

    const escapeHtml = (value) => String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll(String.fromCharCode(34), "&quot;")
      .replaceAll(String.fromCharCode(39), "&#039;")
      .replaceAll("\\n", "<br>");

    const uniqueInOrder = (items) => {
      const seen = new Set();
      const out = [];
      for (const item of items) {
        if (!item || seen.has(item)) continue;
        seen.add(item);
        out.push(item);
      }
      return out;
    };

    const makeOptions = (select, items, emptyLabel) => {
      select.innerHTML = [
        `<option value="">${escapeHtml(emptyLabel)}</option>`,
        ...items.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
      ].join("");
    };

    const renderSource = () => {
      const note = meta.requested_note
        ? `<p>${escapeHtml(meta.requested_note)}</p>`
        : "";
      sourceNote.innerHTML = `
        <strong>참고 파일</strong> ${escapeHtml(meta.source_file)}
        <span> / ${escapeHtml(meta.source_sheet)}, ${escapeHtml(meta.supplement_sheet)}</span>
        ${note}
      `;
    };

    const renderValueTable = (values) => {
      if (!values || values.length === 0) {
        return `<div class="empty">-</div>`;
      }

      const rows = values.map((item) => `
        <tr>
          <td class="value-code" data-label="변수값">${escapeHtml(item.value || "-")}</td>
          <td class="value-label" data-label="변수값설명">${escapeHtml(item.label || "-")}</td>
        </tr>
      `).join("");

      return `
        <div class="values-wrap">
          <table>
            <thead>
              <tr>
                <th>변수값</th>
                <th>변수값설명</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      `;
    };

    const renderCard = (item, isFocused) => {
      const notes = item.notes && item.notes.length
        ? `<div class="note"><strong>비고</strong><br>${item.notes.map(escapeHtml).join("<br>")}</div>`
        : "";

      const businessValues = item.business_values || [];
      const businessBlock = businessValues.length
        ? `
          <section class="extra-values">
            <div class="section-title">사업장업종세분류 상세 코드 <b>${businessValues.length.toLocaleString("ko-KR")}개</b></div>
            ${renderValueTable(businessValues)}
          </section>
        `
        : "";

      return `
        <article class="card ${isFocused ? "is-focused" : ""}">
          <div class="card-head">
            <div>
              <div class="kicker">${escapeHtml(item.table)}</div>
              <div class="var-line">
                <div class="var-name">${escapeHtml(item.variable)}</div>
                <div class="seq">순번 ${escapeHtml(item.seq)}</div>
              </div>
              <div class="desc">${escapeHtml(item.description)}</div>
            </div>
            ${notes}
          </div>
          <div class="card-body">
            <section>
              <div class="section-title">변수 세부설명 <b>${(item.values || []).length.toLocaleString("ko-KR")}개</b></div>
              ${renderValueTable(item.values || [])}
            </section>
            ${businessBlock}
          </div>
        </article>
      `;
    };

    const refreshVariableOptions = () => {
      const selectedTable = tableFilter.value;
      const currentVariable = variableFilter.value;
      const scoped = selectedTable
        ? variables.filter((item) => item.table === selectedTable)
        : variables;

      const names = uniqueInOrder(scoped.map((item) => item.variable));
      makeOptions(variableFilter, names, "전체 변수");

      if (names.includes(currentVariable)) {
        variableFilter.value = currentVariable;
      }
    };

    const render = () => {
      const selectedTable = tableFilter.value;
      const selectedVariable = variableFilter.value;
      const filtered = variables.filter((item) => {
        if (selectedTable && item.table !== selectedTable) return false;
        if (selectedVariable && item.variable !== selectedVariable) return false;
        return true;
      });

      const isFocused = Boolean(selectedTable || selectedVariable);
      cards.innerHTML = filtered.map((item) => renderCard(item, isFocused)).join("");
      visibleCount.textContent = filtered.length.toLocaleString("ko-KR");
      noResults.classList.toggle("is-visible", filtered.length === 0);
    };

    const applyTheme = (theme) => {
      if (theme === "auto") {
        document.documentElement.removeAttribute("data-theme");
      } else {
        document.documentElement.setAttribute("data-theme", theme);
      }
      localStorage.setItem("dataset-codebooks-theme", theme);
    };

    const init = () => {
      totalCount.textContent = variables.length.toLocaleString("ko-KR");
      renderSource();
      makeOptions(tableFilter, uniqueInOrder(variables.map((item) => item.table)), "전체 테이블");
      refreshVariableOptions();

      const savedTheme = localStorage.getItem("dataset-codebooks-theme") || "auto";
      themeSelect.value = savedTheme;
      applyTheme(savedTheme);

      tableFilter.addEventListener("change", () => {
        refreshVariableOptions();
        render();
      });

      variableFilter.addEventListener("change", render);
      themeSelect.addEventListener("change", () => applyTheme(themeSelect.value));

      render();
    };

    init();
  </script>
</body>
</html>
'

html <- paste0(html_head, payload_json, html_tail)

for (out_file in out_files) {
  dir.create(dirname(out_file), showWarnings = FALSE, recursive = TRUE)
  writeLines(html, out_file, useBytes = TRUE)
}

cat("Wrote:\n")
cat(paste0("- ", out_files, collapse = "\n"), "\n")
cat("Variables:", length(variables), "\n")
