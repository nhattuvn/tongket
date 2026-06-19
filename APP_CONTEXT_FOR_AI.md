# Tong Ket Manager - AI Context

This file summarizes the local app so another AI/coding agent can quickly understand and continue the work.

## Purpose

`Tong Ket Manager` is a local Streamlit app for managing freelance project summary data from Excel files stored under:

`C:\Users\Nhat Tu\GEMINI\TONG KET`

The app imports project rows into a local SQLite database, displays them by company/client and period, supports search/edit/manual entry, imports embedded Excel images, and exports summaries to PDF.

## Main App Location

- App folder: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP`
- Main code: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\app.py`
- Database: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\tong_ket.db`
- Uploaded/imported images: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\uploads`
- PDF exports: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\exports`
- Requirements: `C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\requirements.txt`

Run:

```powershell
cd "C:\Users\Nhat Tu\GEMINI\TONG_KET_APP"
streamlit run app.py
```

Current local URL when running:

`http://localhost:8501`

## Current Requirements

`requirements.txt` contains:

```text
streamlit
pandas
openpyxl
reportlab
Pillow
```

## Data Model

The SQLite table is `entries`.

Important columns:

- `id`: local row id.
- `source_key`: unique import key, normally `relative_excel_path|sheet_name|row_number`. It is relative to the selected import folder, not an absolute Windows path.
- `client_name`: company/client grouping, for example `ERNEST`, `ETHAN`, `JOSELYN`, `KING'S CARPENTRY`.
- `project_name`: parsed project title.
- `owner`: parsed owner when project title has ` - OWNER`.
- `period_label`: period such as `JANUARY + FEBRUARY - 2026`.
- `period_year`: parsed year.
- `description`: multiline drawing/view description.
- `drawing_qty`, `unit_price`, `amount`: numeric billing values.
- `status`: `Imported`, `Edited`, or `Manual`.
- `notes`: user notes.
- `image_path`: one or more image paths separated by `|`.
- `source_file`, `source_sheet`, `source_row`: trace back to Excel.
- `deleted_at`: soft delete marker.
- `created_at`, `updated_at`: timestamps.

Project names, owner names, and company/client names are normalized/displayed in uppercase. Use `uppercase_project_name()` for project names and `uppercase_label()` for owners/companies when saving, rendering, searching labels, filtering, or exporting. Existing mixed-case DB rows do not need a destructive migration because the display/export/filter layer uppercases them.

There is also a `settings` table for values such as:

- `payment_info`
- `import_folder`

## Company / Client Handling

Periods can overlap between companies, so the app must group by:

`client_name + period_label`

Do not use `period_label` alone as a unique period.

The importer detects `client_name` using:

1. Top folder under `TONG KET`, unless the folder is generic like `CARPENTRY`, `OLD`, or `HOA DON`.
2. Otherwise, prefix in filename before a dash, such as `Ernest - JANUARY - 2024.xlsx`.
3. Some legacy names like `ERNEST`, `RYAN`, `KELVIN`.
4. Fallback: `GENERAL`.

Examples:

- `TONG KET\CARPENTRY\...\Ernest - JANUARY - 2024.xlsx` -> `ERNEST`
- `TONG KET\ETHAN\8-2024\ETHAN - 2024.xlsx` -> `ETHAN`
- `TONG KET\JOSELYN\02-2026\JOSELYN- FEBRUARY - 2026.xlsx` -> `JOSELYN`
- `TONG KET\KING'S CARPENTRY\...\KING - JULY - 2024.xlsx` -> `KING'S CARPENTRY`

## Period Handling

The app no longer filters out years before 2020 in `load_entries`.

Period detection uses:

1. A period row in Excel, usually column A containing a year and column B empty.
2. Fallback from file/folder name, including month names or numeric folders such as `8-2024`.
3. Excel datetime period cells are normalized to month/year labels.

Important: Some templates place the period before the table header, so parser must preserve `current_period` before and after header detection.

## Excel Import

Main import functions in `app.py`:

- `discover_excel_rows(root)`
- `import_excel_data(root=None)`
- `detect_client_name(path, root)`
- `detect_period_from_path(path, root)`
- `extract_sheet_images(path, sheet, client_name)`

The importer skips:

- Temporary files starting with `~$`.
- `TONG_KET_MASTER_INDEX.xlsx`.
- Total/payment/banking rows.

Project row detection does not require a numeric `No` in column A anymore. Some Excel templates leave `No` blank after the first row. A row is imported when the `Projects` cell has text and either `drawing_qty` or `amount` is numeric/non-zero, while total/payment/banking rows are still skipped.

Supported headers:

- `Projects` + `Drawings Quantity`
- `Projects` + `Views Quantity`
- Generic quantity wording is also accepted.

The importer uses `openpyxl.load_workbook(..., read_only=False, data_only=True)` because embedded images are not available in read-only mode.

## Image Import

Embedded Excel images are extracted from worksheet `_images`.

They are saved under:

`C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\uploads\excel_imports\{client}\{workbook_stem}\...jpg`

Images are mapped to project rows by their Excel anchor row. Multiple images for one project are stored in `image_path` separated by `|`.

Manual uploads are saved under:

`C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\uploads`

Rendering functions:

- `resolve_image_path`
- `image_data_uri`
- `image_paths_from_value`
- `render_image_thumbnails`

PDF export uses ReportLab images from `image_path`.

## Import Update Behavior

Imported rows are keyed by `source_key`.

On re-import:

- Existing absolute-path source keys are migrated to relative source keys before comparing current import rows.
- Imported rows are updated from Excel.
- For `Imported` rows, `image_path` is refreshed from Excel.
- For `Edited` or `Manual` rows, manually edited/uploaded images are preserved.
- Imported rows missing from the current import are soft-deleted by setting `deleted_at`.

Manual entries use `source_key` starting with `manual|`.

The latest import summary is stored in `settings.last_import_summary` as JSON. It includes:

- `imported_at`
- `import_folder`
- `files_found`
- `rows_discovered`
- `rows_upserted`
- `rows_deleted`
- `source_keys_migrated`
- `errors`

## UI Structure

Streamlit tabs after the UI cleanup, ordered by workflow:

- `Tổng quan`
- `Thêm mới`
- `Kỳ thanh toán`
- `Tìm kiếm`
- `Xuất PDF`
- `Cài đặt`

Important UI behavior:

- Global CSS is injected by `inject_app_css()` at the start of `main()`. The app now intentionally uses a fixed light theme, not Streamlit Light/Dark variables. The palette is defined once in `:root`: `--bg`, `--surface`, `--border`, `--text`, `--muted`, `--accent`, `--primary`, and `--soft-accent`.
- UI CSS should use the fixed light palette variables rather than ad-hoc colors. `render_period_table()` is a Streamlit component iframe, so it defines/copies the same fixed variables inside the component HTML. PDF/Excel export code still uses fixed ReportLab/openpyxl colors.
- `Dashboard` uses a Bento-style layout from `display_bento_dashboard()`: 4 metric cards, a 3-column revenue chart card built with HTML/CSS (no Plotly dependency), a top projects card, and a full-width recent-period table section.
- The old standalone `Sua / Xoa` tab was removed.
- `Periods` tab selects `Company` first, then year, then period.
- `Periods` keeps the rich HTML period table in view mode. Clicking `Chỉnh sửa` toggles the same area into inline edit mode, where every row becomes editable at once with project/owner/description/qty/unit/image fields. Row `X` marks a row for deletion, and `Lưu tất cả thay đổi` applies updates plus soft-deletes marked rows. Inline edit mode is wrapped in a bordered frame with a summary strip.
- The Periods table renderer was redesigned in `render_period_table()`: project name and owner/start metadata are separated, descriptions use muted dot bullets, image thumbnails use an enlarged 88px 3-column square grid with `+N` overflow, amount uses the fixed amber `--accent`, and the total row uses a light top border. The table header uses a dark fixed-light header (`--primary` background with `--surface` text), project names are emphasized at 16px/700, description bullets are 13px, and the amount column has amber side borders so row lines remain visible. Image is the last column after Amount. Clicking a Periods thumbnail opens an in-table modal with the original image when an original exists; older imported images fall back to the stored image. `started in...` lines in `description` are extracted by `split_period_project_text()` and rendered as metadata: `Project · Owner (started in PERIOD)`; those lines are omitted from visible description bullets. PDF and Excel export use the same split/format.
- Search/Edit/Export include `Company` filters.
- `Tìm kiếm` groups results by company and period for viewing, but export selection uses a single `st.multiselect` (`search_export_multiselect`) instead of nested checkboxes. This avoids Streamlit rerun/state conflicts. Export buttons are always rendered and disabled until at least one row is selected. Search exports use English titles from `search_export_title()`, for example `ERNEST - Search Summary - "44 Pollen"`, and call export functions with `include_period=True` so mixed-period results include a `Period` column. Search PDF export calls `save_pdf_file(..., include_images=False)` so it omits the `Image` column; Search Excel export keeps images. When search rows are selected, the app also renders a `Search Summary Preview` using `render_period_table(selected_data)`, so the in-app preview still shows image thumbnails.
- `Them moi` was refactored into a two-level layout: top cards for Company/Year/Period, a realtime preview strip above custom project rows, row-level delete buttons, `+ Them dong`, and a summary panel. The old `st.data_editor(num_rows="dynamic")` flow is no longer used there.
- `Thêm mới` initializes with one project row by default; users add more rows manually with `+ Thêm dòng`.
- The project-row entry area in `Thêm mới` is wrapped in a bordered frame with a summary strip.
- In `Thêm mới`, each project row uses the main `Dự án` text input as lightweight autocomplete. Typing a project name shows matching old projects for the selected company as buttons below the field; clicking one auto-fills the project name, latest owner, latest unit price, and stores the first period as row metadata. The UI shows it as `Project · Owner (started in PERIOD)`. On save, the app appends `started in PERIOD` into `description` for persistence without changing the DB schema. If the user only types a name and does not click a suggestion, it is treated as a new project. The helper functions are `project_history_options()`, `project_history_summary()`, and `apply_project_history_to_add_row()`.
- Each `Thêm mới` project row also has `Xem lịch sử`, which shows previous periods for that project in the same company with period, drawing quantity, amount, and a compact description preview.
- If the same project continues with a different unit price, enter it as a separate row rather than trying to merge prices into one project row.
- In `Them moi`, each project row has its own multi-file image uploader. Uploaded images are resized/saved immediately, previewed in that row, and submitted as `image_path` using the existing `path|path` format.
- After saving from `Them moi`, the app reruns once, shows a success message, and renders the just-saved period table at the top of the add tab.
- New manual uploads are saved in two versions: originals under `uploads/originals/` and thumbnails under `uploads/thumbs/`. The database stores thumbnail paths. `original_image_path_for_thumb()` maps a thumb path back to the matching original when possible.
- `Cai dat` manages fixed Company and Owner lists. Values are stored as JSON in `settings.company_owner_config`.
- Payment information is now a section inside `Cài đặt`, not a separate top-level tab.
- Visible UI text is Vietnamese with accents where practical. Internal keys, DB fields, imported values, and business data remain unchanged.
- Company in `Them moi` uses the configured company dropdown. Owner per row uses the configured owner list for the selected company.
- `Tìm kiếm` uses a lightweight autocomplete flow: `Tìm nhanh` is a text input, and matching suggestions for project/company/owner/period appear as buttons under the filters. Clicking a suggestion applies that value as the active keyword for the grouped search results.
- Opening a period from search stores both `search_open_period` and `search_open_client`.
- PDF period export title uses `{client} - {period}`.

## PDF Export

PDF functions:

- `make_pdf(data, title, include_period=False, include_images=True)`
- `save_pdf_file(data, title, include_period=False, include_images=True)`

PDF includes:

- Title.
- Total drawings and amount.
- Table with project, qty, unit, amount, and image as the final column when `include_images=True`. Images are arranged as a 3-column grid inside the PDF image cell. `pdf_image_grid()` currently allows up to 12 images per row. `make_pdf()` uses `calc_pdf_row_height()` to set dynamic ReportLab `rowHeights` from image count and text length so image-heavy rows do not overlap following rows. Project heading is bold in PDF and description lines are italic.
- When `include_period=True`, used by Search export, the table includes a `Period` column so rows from different periods remain readable.
- Payment information from settings.

PDF files are saved in:

`C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\exports`

## Excel Export

Periods can also be exported to Excel from the `Periods` tab using `Xuat Excel`.

Excel export function:

- `save_excel_file(data, title, include_period=False)`

Excel output is saved in:

`C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\exports`

Excel column order matches the current Periods/PDF order:

`No | Project | Drawings Quantity | Unit Price (SGD) | Amount (SGD) | Image`

When `include_period=True`, used by Search export, Excel column order is:

`No | Period | Project | Drawings Quantity | Unit Price (SGD) | Amount (SGD) | Image`

The Image area is at the far right and uses three adjacent sub-columns to visually create a 3-column image grid. Thumbnails are embedded to keep the workbook size reasonable.

Excel export includes a title row using the passed title such as `ERNEST - JANUARY + FEBRUARY - 2026`, summary row, bordered table, payment information, landscape A4 print setup, repeated table header, and fit-to-width printing. Project cells use rich text: the project/owner line is bold and the description lines are italic.

## Current Verified Import State

After the latest import:

- Imported rows discovered from Excel: `184`
- Active rows in database: `188`
- Rows with images: `173`
- Rows missing period: `0`
- Active rows before 2020: `0`

Company counts:

- `ERNEST`: 154 rows, 148 with images
- `ETHAN`: 7 rows, 6 with images
- `JOSELYN`: 13 rows, 6 with images
- `KING'S CARPENTRY`: 10 rows, 9 with images
- `General`: 4 rows, 4 with images

`General` currently appears to be old manual app data, not from the latest Excel import.

Known overlapping periods currently handled correctly by company grouping:

- `AUGUST - 2024`: `ETHAN`, `KING'S CARPENTRY`
- `NOVEMBER - 2025`: `JOSELYN`, `KING'S CARPENTRY`

## Backup

A database backup was created before schema/import changes:

`C:\Users\Nhat Tu\GEMINI\TONG_KET_APP\tong_ket.backup_20260523_091039.db`

## Things Future Agents Should Be Careful About

- Do not group periods only by `period_label`; always include `client_name`.
- Do not restore the old `COALESCE(period_year, 0) >= 2020` filter.
- Do not switch Excel loading back to `read_only=True` if image import is needed.
- Do not overwrite manual or edited images during re-import unless explicitly requested.
- Preserve `|` as the separator for multiple image paths.
- Do not add dark-mode handling or `@media (prefers-color-scheme: dark)` for app UI. The app is intentionally fixed-light.
- Avoid nested checkbox state for search export selection. Use the single multiselect flow unless there is a strong reason to reintroduce checkbox callbacks.
- Apply `split_period_project_text()` consistently in `render_period_table()`, `make_pdf()`, and `save_excel_file()`. Do not render raw `description` lines containing `started in...`.
- Be careful with payment/banking rows in variant Excel templates.
- Some Excel files have no embedded images, so missing images are not always an import bug.
- This is not currently a git repository, so use direct file inspection/backups instead of relying on git history.

## Last Known App Check

The app was started successfully and responded with HTTP status `200` at:

`http://localhost:8501`
