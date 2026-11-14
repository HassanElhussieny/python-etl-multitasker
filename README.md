# sales_etl — Daily Sales & Returns ETL

Processes a folder of ERP CSV exports (Arabic columns) using **one** user-supplied
*Reference Yesterday Date* and produces an Excel report (three total columns + the five business columns) with a Grand Total,
plus a verifiable reconciliation audit.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Run

Put the ERP CSV files directly in the project's `input` folder, then run
`python -m sales_etl` (or the VS Code Run button on `sales_etl/__main__.py`). A menu appears:

```
Available Tasks:
1. Daily
2. Period
3. Customer
4. Representative
5. Sector & Representative Sales Analysis
6. صافي مبيعات العملاء الشهري
7. عدد العملاء الشهري لكل ملف
```

Type the task number. Only the chosen task runs.

**Task 2** asks for the start date, then the end date (e.g. `07-10-2026` and `12-10-2026`). It sums
`صافى` for every row with start ≤ `التاريخ` ≤ end (both days included) and shows, per file and in
total, **إجمالي المبيعات**, **إجمالي المرتجعات** (sign kept) and **الصافي** (= sales + returns).
The result is saved to `output\Period_<execution date>\period_totals_<start>_to_<end>.xlsx`.
After the report is saved and re-read, each file is filed away with the same lock, journal and
safe-move code as task 1:

- counted → `processed\Period_<execution date>\` (same stamp as the report folder)
- unreadable or with rejected rows → `failures\`
- report not verified, or the file changed during the run → stays in `input`

Every CSV in `input` is counted, including content an earlier run already counted (see
"No duplicate handling" below). Task 2 keeps its own journal (`state\period_journal.jsonl`).
`--no-move` keeps the old read-only behaviour. Without the menu:
`python -m sales_etl --start-date 07-10-2026 --end-date 12-10-2026`.

**Task 3 (Customer)** asks no dates and uses every transaction in the input CSV files. It groups by
customer code (column F, `العميل`) and gives the name from column G. It sums `صافى` (column T, sign
kept) by `نوع المستند` (column B):

- **إجمالي المبيعات**: the مبيعات rows
- **إجمالي المرتجعات**: the مرتجعات rows
- **الصافي**: sales + returns

The report holds exactly five columns: كود العميل، اسم العميل، إجمالي المبيعات، إجمالي المرتجعات،
الصافي. It has one row per customer and is saved to `output\Customer_<execution date>\customer_totals.xlsx`.
Without the menu: `python -m sales_etl --task 3`.

**Task 4 (Representative)** works the same way but groups by sales representative. It uses the
code from column I (`المندوب`) and the name from column J; the customer columns F and G are not used.
The report has exactly five columns: كود المندوب، اسم المندوب، إجمالي المبيعات، إجمالي المرتجعات، الصافي.
It has one row per representative code, sorted by code, with no Grand Total row. It is saved to
`output\Representative_<execution date>\representative_totals.xlsx`.
Without the menu: `python -m sales_etl --task 4`.

For tasks 3 and 4, all CSV files in `input` are combined into one report. If one code appears with
different names, it still gets one row: the report uses the first name found (files are read in name
order) and the terminal prints a `WARNING` that lists the other names.

**Files after tasks 3 and 4:** the report is saved and then re-opened and checked cell by cell. Only
after that are the cleanly read CSV files moved to `processed\Customer_<stamp>\` or
`processed\Representative_<stamp>\` (same stamp as the report folder). Moves use the same lock,
journal (`state\customer_journal.jsonl` / `state\representative_journal.jsonl`) and safe,
never-overwriting move as task 1.

- A file that could not be read, or that had rejected rows, stays in `input`.
- If the report cannot be saved or verified, every file stays in `input`.
- Only `*.csv` files directly in `input` are touched.

A file supplied again is counted again (see "No duplicate handling" below). `--no-move` gives
read-only behaviour.

### Task 5: Sector & Representative Sales Analysis

Classifies every transaction through the **Main Data** workbook, by default the single `.xlsx` in
the project's `data` folder other than `fixed code.xlsx` (`--main-data <path>` chooses another). Its
columns are:

- A: كود العميل
- B: اسم العميل
- C: قطاعات العملاء
- D: المندوب

The workbook is only read, never written. Each transaction's customer code (CSV column F) is looked
up in column A. The sector comes from C and the representative from D. The CSV representative
columns (I, J) and the customer name are never used for matching. Sector and representative names
are matched after normalising spaces and spelling variants (e.g. `'الجملة '` = `الجملة`).

After choosing task 5 the program asks `Select report type:`.

1. **Yesterday & Previous Days** asks `Enter yesterday's date (DD/MM/YYYY):`. That date, never the
   computer's date, is the reference:
   - مبيعات / مرتجعات أمس: التاريخ = the date
   - الأيام السابقة: earlier dates
   - later dates: excluded, and counted in the terminal summary

   Columns: مبيعات أمس، مبيعات الأيام السابقة، إجمالي المبيعات، مرتجعات أمس، مرتجعات الأيام السابقة،
   إجمالي المرتجعات، الصافي. In the saved sheets the headers carry the dates: for the date
   05/11/2026 they read `مبيعات يوم 05/11/2026`, `مبيعات حتى 04/11/2026`, `مرتجعات يوم 05/11/2026`,
   `مرتجعات حتى 04/11/2026`, and likewise `عدد عملاء يوم …` / `عدد عملاء حتى …` (one helper,
   `date_headers`, builds them; the day before is the previous calendar day, across month and year
   ends). إجمالي المبيعات / المرتجعات = the day + the earlier days, so they include the date itself,
   once.
2. **All Transactions** asks no date. Columns: إجمالي المبيعات، إجمالي المرتجعات، الصافي.

Amounts come from `صافى` (T), summed in Decimal with their original sign. `نوع المستند` (B) decides
مبيعات / مرتجعات, and any other type is not counted. إجمالي المبيعات / المرتجعات are أمس + الأيام
السابقة, and الصافي = إجمالي المبيعات + إجمالي المرتجعات.

**Distinct customer counts.** After the money columns, the sector and representative sheets show
how many different customer codes each row served:

- Yesterday & Previous Days: عدد عملاء أمس، عدد عملاء الأيام السابقة، إجمالي العملاء المميزين،
  عدد العملاء من الـ224 كود، عدد العملاء خارج الـ224 كود، عدد عملاء المرتجعات فقط
- All Transactions: إجمالي العملاء المميزين، عدد العملاء من الـ224 كود، عدد العملاء خارج الـ224 كود،
  عدد عملاء المرتجعات فقط

A customer is counted only if it has **at least one مبيعات row** in that row and date group. A
customer with sales and returns counts once; a customer with **only مرتجعات is never counted**
(the returns still count in the amounts). A zero-value sale still qualifies. A customer is counted
once per row and date group, however many transactions it has. Only the rows counted in the
amounts are used (same dates), the code (column F) is the identifier (names are ignored), and
blank codes are not counted. إجمالي العملاء المميزين is a
distinct count over the whole period, so a customer in both أمس and الأيام السابقة counts once (it is
not أمس + الأيام السابقة). In `الإجمالي العام` the counts are distinct customers across the whole
report, computed from all transactions directly, not added up from the rows. The third sheet has
no count column.

**The 224-code split.** `fixed code.xlsx`, in the same folder as the Main Data workbook, lists
customer codes in column A of its first sheet (blank cells and repeated codes are ignored; it is
only read, never changed). The two last columns split the customers of إجمالي العملاء المميزين:
those whose code is in the list and the others, so in each row (and in `الإجمالي العام`, again over
the whole report) عدد العملاء من الـ224 كود + عدد العملاء خارج الـ224 كود = إجمالي العملاء المميزين.
Registered and unregistered customers are both split. Codes are compared exactly as the customer
counts identify them (column F; `0100418` = `100418`).

**Returns-only customers.** The last column, عدد عملاء المرتجعات فقط, counts the distinct customer
codes with at least one مرتجعات row and **no** مبيعات row in the report's scope (أمس + الأيام السابقة,
or every transaction). A customer with sales, even a zero-value one, is never returns-only; several
return invoices count once; blank codes are not counted. In `الإجمالي العام` it is the distinct
count over the whole report (a customer with returns in one file and a sale in another is not
returns-only).

**Master Data** (in the second workbook) lists every transaction counted in the amounts, one row each, in
file and line order (ready for filters and Pivot Tables): التاريخ، المندوب، القطاع، كود العميل، اسم العميل،
المبيعات، الصافي، المرتجعات، حالة كود العميل; then `الإجمالي العام` (outside the filter, no serial
number): the distinct representatives under المندوب (real names: not `غير مسجل` / `غير محدد`), the
distinct customer codes under كود العميل (every code, returns-only ones included; blank codes not
counted), and the sums of المبيعات، الصافي، المرتجعات (net = sales + returns). Each representative
sheet has the same الإجمالي العام of its own rows, also when it has none. المندوب and القطاع come from Main Data by the
customer code (`غير مسجل` for a code that is not there or is blank, `غير محدد` for a blank Main Data
cell); كود العميل is the normalised code; اسم العميل comes from the CSV column after the code (G);
الصافي is `صافى` with its sign; المبيعات is the same amount for a sale and 0 for a return, المرتجعات the
same amount for a return (sign kept) and 0 for a sale, so the column totals give net = sales + returns; حالة كود العميل is `داخل قائمة الـ224 كود` or `خارج قائمة الـ224 كود`
from `fixed code.xlsx`. Rejected rows, other document types and (in Yesterday & Previous Days) dates
after the reference date are not in it, exactly as in the amounts.

Each run writes **two workbooks** in `output\Task5_<execution date>\`, built from the same
transactions and the same period. All labels are Arabic, and the layout is right-to-left, with a
frozen navy header, filters, `#,##0.00` amounts and `DD/MM/YYYY` dates.

Every sheet uses bold Segoe UI, centred both ways with wrapped text, and row heights and column widths
that show the text completely (headers 12 pt in a taller first row, body 11 pt; Cover 14 / 13 pt with
taller rows and a prominent total row). Printing: A4 landscape, centred, narrow margins, all columns
on one page width, rows 1-2 (the date line and the header) repeated on every page, a page number in
the footer. Cover and
summary sheets up to 25 rows fit on one page; Master Data, the representative sheets and longer
lists continue onto further pages instead of shrinking the text.

Every sheet starts with the line `تاريخ التقرير: DD/MM/YYYY` (row 1, across the table, aligned to the
left - the top left of the printed page): the reference date, or the run date for All Transactions.
The table starts on row 2.

Every table starts with the column **م**: 1, 2, 3, ... on the data rows (continuing across the Cover
groups), blank on the subtotal and `الإجمالي العام` rows.

**Excluded representative.** Task 5 leaves `دينا كرم ابراهيم` out of both workbooks: the transactions
of the customers Main Data gives her are not counted anywhere (no row, no sheet, not in any total),
and the run checks that her name appears in no saved cell. The CSV files and Main Data are not
changed.

`تقرير تحليل المبيعات والمرتجعات.xlsx`, exactly three sheets. Its print design is black and white (grey shades only, no blue, no dark fill): medium grey header with black bold text, white / very light grey rows, black text and borders, bold medium-grey subtotals with strong borders, `الإجمالي العام` a little darker with a double top border, Segoe UI, black title line. Column widths are fitted to the content (headers wrap) so the table fills the page width at a readable scale. Print: A4 landscape, one page wide, one page high up to 25 rows (longer lists continue onto further pages instead of shrinking the text), minimal margins (0.2" left / right, 0.3" top, 0.35" bottom), centred horizontally and vertically, rows 1-2 (title and header) repeated on every page, print area = the table only. Excel's own black-and-white print mode stays off on purpose: it would drop the grey shading. On `تقرير المندوبين` (one page) the row heights are computed from the page: at the scale set by the width, the free printable height is shared equally by the data, subtotal and total rows (each at most 1.8 times its own height; title and header unchanged), so the table fills the A4 landscape page without a large empty band under it.

| Sheet | Rows |
|---|---|
| تقرير القطاعات | the 8 sectors (always, even at zero), then `عملاء غير مسجلين` (only when shown, see below), then `الإجمالي العام` |
| تقرير المندوبين | the 16 representatives and `د هاني` (always; not the excluded one) in the Cover order: group one then `د هاني` then `إجمالي المباشر`, group two then `إجمالي الوكلاء`, then any other row with transactions (a name Main Data adds, `غير محدد`) and `عملاء غير مسجلين` (only when shown) with `إجمالي المجموعة الثالثة` - the third group only when it has a row, then `الإجمالي العام`. A subtotal sums its rows' amounts and counts the distinct customers of its group |
| العملاء غير المسجلين | كود العميل، اسم العميل (from CSV column G)، إجمالي المبيعات، إجمالي المرتجعات، الصافي: one row per customer code missing from Main Data, then `الإجمالي العام` |

`Representatives_Analysis_<execution date>.xlsx`:

| Sheet | Content |
|---|---|
| Cover | one row per row of تقرير المندوبين (`غير مسجل` / `غير محدد` as in Master Data) plus `د هاني`, in three groups, each followed by its subtotal (see below), then `الإجمالي العام`: م، المندوب، إجمالي البيع، إجمالي المرتجع، صافي، هدف العملاء، عدد العملاء داخل قائمة الـ224 كود، عدد العملاء خارج قائمة الـ224 كود، نسبة تحقيق هدف العملاء |
| Master Data | one row per counted transaction (م + nine columns), then `الإجمالي العام` |
| one sheet per representative | in Cover order: that representative's Master Data rows, same columns and order, then its `الإجمالي العام` (only the total when it has no rows) |
| Pivot for Last Day | the last sheet: the selected reference date only (never today's date, never other days), one row per representative and customer (Cover order, customers by name; the customer name is the CSV customer name, column G, of the customer's first row with one - never Main Data; the code when there is none): اسم المندوب، العميل، عدد الفواتير، عدد العملاء، إجمالي المبيعات، إجمالي المرتجعات، الصافي (عدد العملاء: blank on a customer row; on `إجمالي <المندوب>` the distinct customer codes of that representative - by code, never by name, each once; on `الإجمالي العام` the distinct customer codes of the whole day, never the sum of the subtotals; the reconciliation recounts them from the raw source rows); after each representative's customers its subtotal `إجمالي <المندوب>` (bold, strong black borders), then `الإجمالي العام` (= the sum of the subtotals). Row 1 shows `تاريخ تشغيل التقرير:` = the RUN date (the computer's date when Task 5 was run, the same moment that names the run folder) and then the transactions' day (the selected reference date) - two different dates. عدد الفواتير = the customer's distinct documents of the day, sales AND returns, identified by `نوع المستند` + `رقم المستند` (a sale and a return with the same number are two documents; the rows of one document, also when it is in several files, count once; a row without a number is a document of its own); a document counts only when at least one of its rows has a non-zero `صافى` (column T) - a document whose rows are all zero is not counted. The count never changes an amount: sales and returns are the sum of every row, zero rows and repeated rows included. A subtotal / الإجمالي العام adds its customers' documents. The reconciliation recounts the documents from the raw source cells (type, number, صافى, date, customer -> Main Data) and names any missing or extra document; returns keep their sign; الصافي = sales + returns. All Transactions has no selected day: the sheet says so and holds only a zero total. Black-and-white design and print setup (A4 landscape, black and white, one page wide, one page high up to 25 rows, rows 1-2 repeated) |

In Cover, the customers inside / outside the 224 list are the customers with a sale (as إجمالي
العملاء المميزين) split by the 224-code list; صافي = إجمالي البيع + إجمالي المرتجع (returns keep their
minus sign: 500,000 + -25,000 = 475,000). In `الإجمالي العام` the amounts are the column sums and the
customer counts are distinct over the whole report.

**هدف العملاء** is fixed (`COVER_TARGETS` in `sales_etl/analysis.py`, never derived from the data):
هند محمد رضا 25، اسلام 25، سامح 38، داليا 44، سميرة 25، محمود الشيمي 21، د هاني 3، ندا 25، هشام 2،
ابراهيم سلام 4، رامي 1، محمد رأفت 2، احمد محمد علام 1، احمد محمود موسي 2، احمد الحريري 1 (المباشر 206،
الوكلاء 13، الإجمالي 219). A row without a target (`اون لاين`, `الاداره`, `غير مسجل`, a new name) shows
`لا يوجد هدف` and `—` as its percentage. **نسبة تحقيق هدف العملاء** = customers inside the 224 list ÷
target (stored as a fraction, shown as `0.00%`: 10 ÷ 38 = 26.32%; no division by zero). For a subtotal
and `الإجمالي العام` it is the sum of the inside customers of the representatives that have a target ÷
the sum of their targets - never a sum or average of percentages. A sheet name keeps
Excel's rules (31 characters at most, no `[ ] : * ? / \`, unique); the full representative name is
always in the cells.

Cover groups, in this exact order:

1. اون لاين، الاداره، هند محمد رضا، اسلام حمد كمال محمد، سامح السيد محمود احم، داليا صلاح الدين فكر،
   سميرة محمد اسماعيل، محمود الشيمي، ندا عبد الناصر طه، د هاني → `إجمالي المباشر`
2. هشام احمد محمد، ابراهيم سلام، رامي رأفت يوقيم، محمد رأفت شحاتة محمد، احمد محمد علام محمد،
   احمد محمود موسي، احمد الحريري → `إجمالي الوكلاء`
3. every other row with transactions (a name Main Data adds, `غير محدد`) and `غير مسجل` (only when
   shown) → `إجمالي المجموعة الثالثة`; no third group and no third subtotal when it has no row

The groups are defined once (`cover_groups`) and used by Cover, تقرير المندوبين, the representative
sheet order and the Pivot order: `د هاني` (`DIRECT_ADDITION`) is always the last row of المباشر, right
above `إجمالي المباشر`, exactly once, also without transactions (then a row of zeros and an empty
sheet). A representative of the two fixed lists is never hidden: one without transactions keeps a row
of zeros with its target, and its target stays in the subtotal and the total.

**Unregistered customers are shown only when they matter.** `عملاء غير مسجلين` (تقرير القطاعات,
تقرير المندوبين), `غير مسجل` (Cover) and the third-group subtotal that would hold only it appear only
when at least one unregistered customer has a transaction with a non-zero `صافى` (`unregistered_shown`,
decided on the transactions, never on a placeholder row). Otherwise none of these rows is written - no
blank row, no zero-only row or subtotal - and the sheet `العملاء غير المسجلين` holds the single line
`لا يوجد عملاء غير مسجلين في بيانات هذا التقرير` instead of a zero total. Unregistered transactions whose
`صافى` are all zero are real rows: they stay in Master Data and in their own `غير مسجل` sheet (and the
Pivot), and `الإجمالي العام` still counts them (its amounts and customers are all transactions, unchanged
either way). A subtotal sums the
group's amounts and targets and counts the **distinct** customers of the group (never a sum of the
rows); `الإجمالي العام` counts the distinct customers of the whole report. Cover is black and white:
grey bold header, black borders, light-grey stripes, subtotals grey with a strong border,
`الإجمالي العام` larger and taller with a double border; A4 landscape, black-and-white printing, one
page wide (one page high up to 25 rows), rows 1-2 repeated.

**Customers missing from Main Data** never stop the run, and their amounts are never dropped.
They are counted once in the `عملاء غير مسجلين` row of both main sheets and listed per code in
the third sheet. Main Data is never changed automatically. Once you add the customer there, the
next run classifies it normally.

If Main Data uses a sector or representative name that is not in the fixed lists, it gets its own
row (before `عملاء غير مسجلين`), and the terminal prints a warning. The run stops before
processing anything if:

- the Main Data workbook is missing or has the wrong headers
- `fixed code.xlsx` is missing, unreadable or has no codes
- the `data` folder holds more than one workbook besides `fixed code.xlsx`
- one customer code appears with two different sectors or representatives

**الإجمالي العام** is the sum of each sheet's detail rows. Before any file moves:

1. Both saved workbooks are re-opened, and their sheet names (exactly Cover, Master Data and one
   sheet per representative in the second), headers, every cell and every grand total are checked;
   every saved Master Data transaction must be in exactly one representative sheet.
2. The reconciliation is run:
   - sector and representative totals are equal
   - the `عملاء غير مسجلين` rows equal the third sheet's total
   - every transaction is counted exactly once
   - the grand totals equal a direct sum of all transactions
   - the grand-total customer counts equal the distinct customers of the rows, and no customer is
     counted in two rows
   - every row's customer counts, and the grand total's, equal a separate plain-Python recount
     from each source row's raw date, document type text and code (customers with مبيعات only),
     including the 224-code split and the returns-only customers
   - in every row, the two split columns add up to إجمالي العملاء المميزين
   - Master Data has exactly the counted transactions (row count = the totals' transactions = the
     source rows filtered again), its sales, returns and net equal the report totals, every sector
     and representative row equals its Master Data rows (amounts and customer counts), and every
     row's representative, sector and 224-code status match Main Data and `fixed code.xlsx`
   - every Cover row equals its تقرير المندوبين row, its Master Data rows (sales, returns,
     صافي = sales + returns), the independent customer recount (inside / outside the 224 list), its
     fixed target and inside ÷ target; its total equals the Master Data totals and the distinct
     customers of the whole report
   - the Cover groups are exactly المباشر (with د هاني) and الوكلاء in order, then the other rows, and
     no representative of تقرير المندوبين is missing; every subtotal equals its group's Master Data rows,
     an independent recount of the group's distinct customers, the sum of its targets and Σinside ÷
     Σtarget
   - column م is 1, 2, 3, ... on every saved table, the excluded representative is in no saved cell
     and none of her transactions is counted (also re-checked from the source rows)

A failure in either workbook names the workbook, sheet and row, and no input file is moved.

**Files:** same lifecycle as tasks 3 and 4 (lock, journal `state\task5_journal.jsonl`, crash
recovery, safe moves).

- Successfully read CSV files move to `processed\Task5_<stamp>\` (same stamp as the report folder)
  only after the checks above pass.
- Failed files and files with rejected rows stay in `input`.
- A save, verification or reconciliation failure moves nothing.
- `--no-move` only reads.

Without the menu:

```powershell
python -m sales_etl --task 5 --ref-date 30/10/2026     # Yesterday & Previous Days
python -m sales_etl --task 5 --report-type 2           # All Transactions
```

### Task 6: صافي مبيعات العملاء الشهري

Asks no date. For every customer code and every month, it sums `صافى` (T) of the مبيعات and
مرتجعات rows, with each sign kept: **net only**, no separate sales or returns. Like task 3, there
is no date filter, and future dates count normally. One year of data is expected; the month is
the month of `التاريخ` (A).

| Column | Content |
|---|---|
| كود العميل | `العميل` (F), exactly as written (no zero stripping, no digit conversion) |
| اسم العميل | the column after the code (G); the first name found (a different name for the same code prints a warning) |
| المندوب | the column after `المندوب` (J): every distinct name for that customer, in order of first appearance by transaction date, separated by `، `. Blank names are skipped, and spelling variants (spaces, أ/ا …) count as one name. |
| يناير … | one column per month, from يناير to the **last month in the data**; months without transactions show 0 |

- One row per customer code. Equal names with different codes stay separate rows, and a customer
  whose net is 0 is kept.
- There is no total row or total column.
- Other document types are not counted; the terminal shows how many there were.
- A مبيعات/مرتجعات row **without a customer code** is rejected like an invalid date or amount: it
  is not counted, the terminal names its file and line, and the file stays in `input`.
- A file without the `العميل` or `المندوب` column fails.

Output: `output\صافي_العملاء_الشهري_<execution date>\تقرير صافي مبيعات العملاء الشهري.xlsx`, one sheet
`التقرير` (right-to-left, navy header, frozen header, `#,##0.00`).

Before any file moves, the workbook is re-opened (sheet, headers including each month, row count,
every cell). Then it is reconciled against an independent sum of the counted rows:

- every customer × month cell;
- the number of customer rows = the number of distinct codes;
- each المندوب cell lists every name exactly once;
- the sum of all cells = the net of all counted rows.

**Files:** the same lifecycle as task 3 (lock, journal `state\سجل_صافي_العملاء_الشهري.jsonl`,
crash recovery, safe moves). Clean files move to `processed\صافي_العملاء_الشهري_<stamp>\`. Failed
files and files with rejected rows stay in `input`, and a save or verification failure moves
nothing. Every file is counted, even if it was processed before. `--no-move` only reads.

```powershell
python -m sales_etl --task 6
```

### Task 7: عدد العملاء الشهري لكل ملف

Asks no date. For **each CSV file on its own**, it counts the distinct customer codes (`العميل`, F)
in every **accounting month** of `التاريخ` (A). Every document type counts, and amounts are never
used. A code that appears many times in one month of one file counts once. It counts again in
another month, and again in another file: files are never merged. Codes are kept exactly as
written, and the customer name is not used. Future dates count normally. One year of data is
expected.

An accounting month (task 7 only) runs from the **26th of the previous month to the 25th** of the
month itself. For example, 26-09-2026 … 25-10-2026 is أكتوبر, and 26-12-2025 … 25-01-2026 is
يناير.

| Column | Content |
|---|---|
| اسم الملف | the CSV file name; one row per file read, with numbers in file names in numeric order (`3omla29` before `3omla210`) |
| يناير … | the number of distinct customer codes of that file in that month, from يناير to the **last month in any file read**; months without customers show 0 |
| إجمالي العملاء المميزين خلال السنة | the last column: the number of distinct customer codes in the **whole file**. A customer in several months counts once, so this is not the sum of the months. An empty file shows 0. |

- A file with no transactions (header only, or no content at all) is **not** a failure: it gets
  a row of zeros and moves like any other clean file.
- A file that cannot be read (no header row, bad encoding, no `العميل` column) fails. It has no
  row and stays in `input`.
- A row with an invalid date or amount, or **without a customer code**, is rejected. The terminal
  names its file, line and reason. The file's other rows are still counted, and the file stays
  in `input`.
- There is no total row or total column.

Output: `output\عملاء_الملفات_الشهري_<execution date>\تقرير عدد العملاء الشهري لكل ملف.xlsx`, one sheet
`التقرير` (right-to-left, navy header, frozen header, `#,##0`).

Before any file moves, the workbook is re-opened (sheet, headers including each month, row count,
every file name and every cell). Then it is reconciled against an independent recount of the
counted rows:

- every file × month cell;
- every year cell = the distinct codes of all counted rows of that file;
- the report rows = the files read, in the same order and with the same names;
- every counted row is in its own file's row.

**Files:** the same lifecycle as task 3, with journal `state\سجل_عملاء_الملفات_الشهري.jsonl`.
Clean files move to `processed\عملاء_الملفات_الشهري_<stamp>\`. Failed files and files with
rejected rows stay in `input`, and a save or verification failure moves nothing. Every file is
counted, even if it was processed before. `--no-move` only reads.

```powershell
python -m sales_etl --task 7
```

**Task 1** is described in the rest of this document. Without the menu:

```powershell
python -m sales_etl --ref-date 18/10/2026
```

If you leave out `--ref-date`, the date is asked for **once** and applied to every file.
Before processing, the date is shown back in an unambiguous form
(`Sunday 18 October 2026 (ISO 2026-10-18)`).

| Option | Meaning |
|---|---|
| `--task 1` … `--task 7` | Choose the task without the menu. |
| `--report-type 1` / `2` | Task 5: Yesterday & Previous Days (needs `--ref-date`) / All Transactions. |
| `--main-data` | Task 5: the Main Data workbook (default: the single `.xlsx` in `<project>\data` other than `fixed code.xlsx`, which is read from the Main Data workbook's folder). |
| `--input-dir` | Another input folder (default: `<project>\input`). |
| `--output-dir` | Where run folders are created (default: `<project>\output`). |
| `--no-move` | Read-only mode: input files are never moved and no journal is kept. |
| `--processed-dir`, `--failures-dir`, `--state-dir` | Override the folders used by automatic mode. By default they sit next to `input`. |
| `--encoding` / `--delimiter` | Force a source encoding / delimiter (`;`, `,`, `tab`). By default both are detected per file. |
| `--recursive` | Also scan sub-folders (only with `--no-move`). |
| `--non-interactive` | Never prompt; missing arguments become an error. |

Exit codes:

| Code | Meaning |
|---|---|
| `0` | `SUCCEEDED`: report generated and verified, audit PASS, every file archived |
| `1` | A report was generated but the audit FAILed or a file failed/was not archived, or no verified report could be produced (`FAILED`) |
| `2` | Invalid arguments, or another run is in progress |
| `3` | No CSV files in `input`: no report was produced |

## Automatic file management

Every `*.csv` directly in `input` is processed in the current run. This includes content that
was processed before, under any name. Nothing is skipped because of file content or processing
history, so supplying the same file again counts its transactions again, by design.

A file is moved only after the report, the audit workbook, the audit JSON and the log have been
written, read back and verified.

| Result | Destination |
|---|---|
| Processed, and none of the file's own checks FAIL | `processed\Run_<execution date>\` (one folder per run) |
| The file itself failed: unreadable, empty, missing column, rejected rows, or an ERP control-total mismatch | `failures\` |
| A system problem: outputs not verified, grand-total inconsistency, or the file changed during the run | Stays in `input`. Nothing is recorded as done. |

Folder names use the **execution date** (the day the program runs), never the reference date:

- `output\Report_2026-10-20\` holds that run's report, audit workbook, audit JSON, log and
  `run_manifest.json`.
- `processed\Run_2026-10-20\` holds the CSV files that run processed successfully.

A second run on the same day uses `Report_2026-10-20_HHMMSS` / `Run_2026-10-20_HHMMSS` (plus
`_2`, `_3`, … if needed). A run's report folder and archive folder always share the same suffix,
and an existing folder is never reused.

Older runs keep their original layout (`output\run_ref-…`, files directly in `processed\`).
Files that older versions placed in `duplicates\` or `state\fingerprints\` are left as they are;
nothing reads or writes those folders any more.

Moving rules:

- Nothing is ever overwritten. If a file of the same name already exists in the destination,
  the new one is saved as `name__<timestamp>_<sha8>.csv`.
- After every move, the destination's SHA-256 is checked against the original.

### No duplicate handling (all five tasks)

No task checks for duplicates:

- No file is skipped, rejected or set aside because it was processed before, has the same name or
  the same content as another file, or appears in a journal. Journals are a crash-safety trace only.
- No transaction row is skipped, removed or flagged because it repeats another row, document
  number, invoice number or customer code. Every valid row is counted by the task's own rules.

So supplying the same file twice, or two copies of it, counts its transactions twice. This is by
design.

### Crash safety (`state\processing_journal.jsonl`)

The journal is a trace only: it never causes a file to be skipped.

It is append-only, and every line is flushed to disk before the run continues. A
lock (`state\etl.lock`) prevents two runs at the same time; the operating system releases it
if the program dies. Each run goes through these steps:

1. `RUN_STARTED` is written.
2. Outputs are written and verified.
3. A single `COMMIT` line records every file's disposition and destination.
4. Each file is moved, and a `MOVED` line is written for it.
5. `RUN_FINISHED` is written.

If a run is interrupted, the next run repairs the situation first:

| Interrupted | What the next run does |
|---|---|
| Before `COMMIT` | The files are still in `input` and are processed normally. The interrupted run folder gets a `RUN_INTERRUPTED.txt` saying its outputs are void. |
| After `COMMIT` | Moves that did complete are confirmed in the journal. A file still in `input` is processed again in the current run, and its old pending move is closed as superseded. |
| A move failed (e.g. the file was open in Excel) | The file stays in `input` and is processed again by the next run. |

`output\Report_…\run_manifest.json` records:

| Field | Meaning |
|---|---|
| `status` | `SUCCEEDED`, `COMPLETED_WITH_PROBLEMS` or `FAILED` |
| `report_generated` | Whether a verified report was produced |
| `execution_date` / `execution_time` | When the program ran |
| `reference_date` | The "yesterday" date that was entered |
| `report_dir`, `processed_run_dir` | This run's two folders |
| `outputs` | Each output's path, and whether it **exists on disk** |
| `files` | For every file: disposition, reason, SHA-256 and current location |

`SUCCEEDED` is used only when the report, audit workbook, audit JSON and log were actually
written and verified. A failed run also gets a manifest, with the error message.

Dates are always **day first**. Accepted forms: `DD/MM/YYYY`, `DD-MM-YYYY`, `DD.MM.YYYY`
and ISO `YYYY-MM-DD`. Two-digit years and impossible dates are rejected.

## Business rules

Only the `صافى` column is summed, with its original sign. Returns stay negative and are
never negated.

| Report column | `نوع المستند` | `التاريخ` |
|---|---|---|
| إجمالي مبيعات أمس | مبيعات | = reference date |
| إجمالي مبيعات الأيام السابقة | مبيعات | < reference date |
| إجمالي مرتجعات أمس | مرتجعات | = reference date |
| إجمالي مرتجعات الأيام السابقة | مرتجعات | < reference date |

- Transactions dated after the reference date are excluded from every metric. They are
  still counted and shown in the audit.
- When no rows match a metric, the value is `0`, never blank or NaN.

The last three report columns count **distinct customer codes** (`العميل`, column F, exactly as
written) that have **at least one valid مبيعات row** dated on or before the reference date. A
customer with sales and returns counts once; a customer with **only مرتجعات is never counted**.
Blank codes, rejected rows, other document types and later dates are not counted, and amounts
don't matter (a zero-value sale still qualifies).

| Report column | Customer codes counted |
|---|---|
| عدد عملاء أمس | with a مبيعات row dated the reference date |
| عدد عملاء الأيام السابقة | with a مبيعات row dated before the reference date (all those days together) |
| إجمالي العملاء المميزين | in either period: the union of the two, not their sum |

A customer is counted once per column however many rows they have. Each file is counted on its
own. The Grand Total row counts the **union** of the files' code sets, never the sum of the file
rows, so a customer in several files counts once. A customer with returns in one file and a sale
in another is counted in the Grand Total; one with only returns in every file is not.

## Outputs

Each run writes to a new folder, `output/Report_<execution date>/` (with a suffix for further
runs on the same day). Existing files are never overwritten, and source CSV contents are never
modified.

| File | Content |
|---|---|
| `sales_report_ref-<date>.xlsx` | Sheet **التقرير**: first three totals - **إجمالي المبيعات** (مبيعات أمس + مبيعات الأيام السابقة), **إجمالي المرتجعات** (مرتجعات أمس + مرتجعات الأيام السابقة, sign kept), **الصافي** (إجمالي المبيعات + إجمالي المرتجعات), computed in Python and written as numbers - then the five columns, then the three customer counts (عدد عملاء أمس، عدد عملاء الأيام السابقة، إجمالي العملاء المميزين; customers with مبيعات only); one row per successfully processed file, then **Grand Total**. Sheet **بيانات التقرير**: the reference date, the rules, file counts, failed files and the audit status. |
| `audit_ref-<date>.xlsx` | Sheets **Summary**, **Checks** (PASS/WARN/FAIL), **Files** (encoding, delimiter, row counts, excluded future rows, ERP control totals, SHA-256), **Issues**, **Transactions** (every row and the metric it went into), **Rejected rows**. |
| `audit_ref-<date>.json` | The same information in machine-readable form. Amounts are exact decimal strings. |
| `pipeline.log.jsonl` | Structured log, one JSON object per line. |

## How files are read

- **Encoding:** detected per file in this order: byte-order mark, then strict UTF-8, then
  cp1256 (Windows Arabic, which the ERP uses).
- **Delimiter:** chosen as the delimiter under which a row containing all required headers
  is found. A comma can't be confused with the delimiter just because amounts contain
  thousands separators.
- **Header row:** located by its content, so the ERP preamble (title, `من … الى …` period,
  branch and warehouse lines) is skipped. Header matching tolerates spelling variants
  (`صافى`/`صافي`, tatweel, `أ`/`ا`). The original column names are preserved.
- **Required columns:** `التاريخ`, `نوع المستند`, `صافى`.
  - If one is missing, or **appears twice**, the file fails. The pipeline won't guess which
    `صافى` column holds the real data.
  - Blank headers get safe placeholder names.
  - Other duplicated headers are renamed and reported.
- **ERP footer rows** (`عدد المستندات`, `الاجمالى العام`) are recognised and used as control totals.

## Validation policy

| Situation | Handling |
|---|---|
| Empty file, unreadable encoding, missing or ambiguous required column | **File-level failure.** The file is left out of the report, listed with its reason, and the batch continues. Audit = FAIL. |
| File with a header but no rows | Processed; all metrics are 0, with a warning. |
| Invalid date, invalid or empty `صافى`, empty `نوع المستند`, row with more cells than the header (possible column shift) | **Row rejected** and logged with its line number and raw content. It is never counted as zero. The file is still processed. Audit = FAIL (`NO_REJECTED_ROWS`). |
| `نوع المستند` other than مبيعات/مرتجعات | Kept in the ledger, excluded from the metrics, warning. |
| Sales with a negative amount, or returns with a positive amount | Kept as-is, warning. |
| Reference date outside the file's `من … الى …` period | Warning. |

Repeated rows, document numbers and whole files are neither detected nor removed: every valid
row is counted (see "No duplicate handling").

## Reconciliation checks

All amounts are `decimal.Decimal`; nothing is computed in floating point.

| Check | What it proves |
|---|---|
| `METRIC …` (×4 per file) | The pandas result equals a separate plain-Python recomputation from each row's raw date and document type. Amount and row count are both compared. |
| `FUTURE_TRANSACTIONS_EXCLUDED` | No row dated after the reference date reached a metric. |
| `PARTITION_COMPLETE` | The 4 metrics + future rows + other document types = all valid rows. Nothing is lost or counted twice. |
| `ERP_CONTROL_TOTAL_NET` / `ERP_CONTROL_DOC_COUNT` | The sum of all rows' `صافى` and the row count match the ERP's own footer totals. |
| `NO_REJECTED_ROWS` | Every transaction row was parsed. |
| `GRAND_TOTAL …` (×4) | The sum of per-file totals equals a recomputation over all source transactions. |
| `CUSTOMER_COUNTS` (per file) / `GRAND_TOTAL_CUSTOMERS` | The three customer counts, computed from code sets, equal a separate plain-Python recount (customers with a مبيعات row only) from each row's raw date, document type and code. The Grand Total recount runs over all files' rows at once. |
| `ALL_FILES_PROCESSED`, `FILES_DISCOVERED` | The batch is complete. |
| `SOURCE_FILES_UNCHANGED` | SHA-256 of every input file is identical before and after the run. |
| `REPORT_READBACK` | The saved workbook is re-opened and every cell matches the computed values. |

The overall audit is **PASS** only if no check is FAIL. WARN and N/A don't fail the audit.

## Tests

```powershell
python -m pytest
```

The tests use small synthetic datasets with known totals, written in the exact ERP layout.

`tests/fixtures/alex8_compact.txt` holds the numeric content of a real export
(`alex8.csv`). The test rebuilds it as a cp1256 file and checks it against that file's
original footer (235 documents, `صافى` 1,105,646.62). Customer names in the fixture are
synthetic.

## Code layout

| Module | Responsibility |
|---|---|
| `config.py` | Column names, document types, report columns, `PipelineConfig` |
| `dates.py`, `amounts.py`, `text.py` | Strict parsing; Arabic-aware matching |
| `ingest.py` | File discovery, encoding/delimiter detection, header location and column resolution |
| `normalize.py` | Row classification and row-level validation |
| `metrics.py` | The four metrics (pandas + Decimal) |
| `reconcile.py` | Independent recomputation, ERP control totals |
| `report.py` | Report workbook, audit workbook and JSON, read-back verification |
| `pipeline.py` | Orchestration and per-file failure isolation |
| `period.py` | Task 2: totals for a chosen period |
| `customer.py` | Task 3: totals per customer (columns F/G) |
| `representative.py` | Task 4: totals per representative (columns I/J) |
| `grouped.py` | Shared code for tasks 3 and 4 (grouping, report, verification); the file lifecycle shared by tasks 3, 4 and 5 (`run_batch_managed`) |
| `analysis.py` | Task 5: Main Data mapping, the report and the Representatives_Analysis workbook, verification and reconciliation |
| `monthly_customer.py` | Task 6: monthly net per customer, verification and reconciliation |
| `monthly_file_customers.py` | Task 7: monthly distinct customers per file, verification and reconciliation |
| `cli.py` | Task menu, prompts and execution summary |
