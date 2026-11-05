from pathlib import Path
from datetime import datetime
from decimal import Decimal
import io
import re
import tkinter as tk
from tkinter import filedialog, messagebox

from pypdf import PdfReader

try:
    import xlsxwriter
except ImportError:
    raise ImportError(
        "This script needs the XlsxWriter package. In a notebook cell, run: "
        "%pip install --upgrade xlsxwriter   (then restart the kernel)"
    )

# Row colors in the spreadsheet.
ROW_COLOR = "#FFFFFF"        # every row
DONE_COLOR = "#C6EFCE"       # row turns this color when Updated AND Reprocessed
HEADER_COLOR = "#002E5D"     # header row

# ---------------------------------------------------------------------------
# Patterns. Everything inside a single line uses [ \t] instead of \s so that
# a match can never wander across a line break (or a page break).
# ---------------------------------------------------------------------------

# First line of every record:
#   [Recipient ID] time time-zone entry-class amount reason-code
# The recipient ID is optional; the amount is ".00" on Notifications of Change.
RECORD_PATTERN = re.compile(
    r"^[ \t]*(?:(?P<recipient_id>\S.*?)[ \t]+)?"
    r"\d{1,2}:\d{2}:\d{2}[ \t]+[AP]M[ \t]+\S+[ \t]+"
    r"[A-Z]{2,4}[ \t]+"
    r"(?P<amount>[\d,]*\.\d{2})[ \t]+"
    r"(?P<reason>[RC]\d{2})[ \t]*$",
    re.MULTILINE,
)

# Second line of every record: name, effective date, account type, and the
# first line of the reason-code description.
NAME_LINE_PATTERN = re.compile(
    r"(?P<name>.*?)[ \t]*(?P<date>\d{2}/\d{2}/\d{4})"
    r"[ \t]+[A-Z]{2,4}\b[ \t]*(?P<first>.*)$"
)

TRACE_PATTERN = re.compile(r"Trace\s+Number\s*:\s*(\d+)\s+(\d+)")
BANK_PATTERN = re.compile(r"Receiving\s+Bank\s+ID\s*:[ \t]*(\d+)")
ACCOUNT_PATTERN = re.compile(r"Receiving\s+Account\s*:[ \t]*([0-9A-Za-z*-]+)")
TRACE_LABEL_PATTERN = re.compile(r"Trace\s+Number\s*:")

PAGE_FOOTER_PATTERN = re.compile(r"^\s*Page:\s*\d+\s*$")
REPORT_DATE_PATTERN = re.compile(r"As\s+of\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)


def clean_page(page_text, page_number):
    """Remove the header and footer that Wells Fargo repeats on every page,
    so a record that continues onto the next page reads as one record."""
    lines = page_text.splitlines()

    if lines and lines[0].strip().startswith("ACH Return/NOC Report"):
        for index, line in enumerate(lines[:15]):
            if "Treasury Information Reporting" in line:
                lines = lines[index + 1:]
                break
        else:
            print(
                f"WARNING: could not find the end of the page header on "
                f"page {page_number}; the header was left in place."
            )

    return "\n".join(
        line for line in lines if not PAGE_FOOTER_PATTERN.match(line)
    )


def read_description(lines, start):
    """Collect a reason-code description that may wrap onto several lines.
    It ends at the 'Credit'/'Debit' line (which also holds the Company Entry
    Description, e.g. 'PAYROLL Credit') or at 'Original'."""
    parts = []

    for line in lines[start:]:
        if re.search(r"(?:^|\s)(?:Credit|Debit)$", line) or line.startswith(
            "Original"
        ):
            break
        if line:
            parts.append(line)

    return " ".join(parts)


def extract_rows(page_texts):
    """Turn the text of every PDF page into
    (columns, report_date, rows, skipped_count)."""
    if not any(text.strip() for text in page_texts):
        raise ValueError(
            "No readable text was found. This PDF may need OCR first."
        )

    # Report date comes from the 'As of' heading.
    report_dates = set(REPORT_DATE_PATTERN.findall("\n".join(page_texts)))

    if len(report_dates) != 1:
        raise ValueError(
            "Expected one report date in the PDF; "
            f"found: {sorted(report_dates)}"
        )

    report_date = datetime.strptime(report_dates.pop(), "%m/%d/%Y")

    # Clean each page, then join them so records can span pages.
    cleaned_pages = [
        clean_page(text, number)
        for number, text in enumerate(page_texts, start=1)
    ]

    for number, text in enumerate(cleaned_pages, start=1):
        if not text.strip():
            print(
                f"WARNING: page {number} has no readable text (it may be "
                "an image). Records on that page would be missing."
            )

    text = "\n".join(cleaned_pages)

    record_starts = list(RECORD_PATTERN.finditer(text))

    if not record_starts:
        raise ValueError(
            "No return or NOC records matched the expected report layout."
        )

    columns = [
        "Employee Name",
        "Email",
        "Original Trace Number",
        "Returned Trace Number",
        "Receiving Bank ID",
        "Receiving Account",
        "Amount",
        "Reason Code",
        "Reason Code Description",
        "Effective Date",
        "Report Date",
    ]

    rows = []

    for index, match in enumerate(record_starts):
        end = (
            record_starts[index + 1].start()
            if index + 1 < len(record_starts)
            else len(text)
        )
        block = text[match.end():end]
        lines = [line.strip() for line in block.split("\n")]

        employee = None

        for position, line in enumerate(lines):
            found = NAME_LINE_PATTERN.match(line)

            if found and found["name"]:
                continued = read_description(lines, position + 1)
                description = " ".join(
                    part for part in (found["first"], continued) if part
                )
                employee = {
                    "name": found["name"],
                    "date": found["date"],
                    "description": description,
                }
                break

        traces = TRACE_PATTERN.search(block)
        bank = BANK_PATTERN.search(block)
        account = ACCOUNT_PATTERN.search(block)

        missing = [
            label
            for label, value in [
                (
                    "employee name / effective date / description",
                    employee,
                ),
                ("original and returned trace numbers", traces),
                ("receiving bank ID", bank),
                ("receiving account", account),
            ]
            if value is None
        ]

        if missing:
            raise ValueError(
                f"Record {index + 1} could not be fully read. "
                f"Missing: {', '.join(missing)}. "
                "No Excel file has been written."
            )

        # Every record must hold exactly one trace-number line.
        if len(TRACE_LABEL_PATTERN.findall(block)) != 1:
            raise ValueError(
                f"Record {index + 1} does not contain exactly one trace "
                "number. No Excel file has been written."
            )

        rows.append([
            employee["name"].strip(),
            None,
            traces.group(1),
            traces.group(2),
            bank.group(1),
            account.group(1),
            float(Decimal(match["amount"].replace(",", "") or "0")),
            match["reason"],
            employee["description"].strip(),
            datetime.strptime(employee["date"], "%m/%d/%Y"),
            report_date,
        ])

    # Check that no trace-number record was silently skipped.
    trace_count = len(TRACE_LABEL_PATTERN.findall(text))

    if trace_count != len(rows):
        raise ValueError(
            f"Found {trace_count} trace-number records but parsed "
            f"{len(rows)} employee rows. "
            "No Excel file has been written."
        )

    # Notifications of Change always have a $0.00 amount and never need
    # follow-up, so only keep records with an amount greater than 0.
    # (They are still parsed above so the trace-number check stays honest.)
    payable_rows = [row for row in rows if row[6] > 0]

    if not payable_rows:
        raise ValueError(
            "No records with an amount greater than 0 were found."
        )

    return columns, report_date, payable_rows, len(rows) - len(payable_rows)


def build_workbook_bytes(columns, rows):
    """Build the Excel file in memory and return its bytes."""
    if not hasattr(xlsxwriter.worksheet.Worksheet, "insert_checkbox"):
        raise ImportError(
            "Your XlsxWriter version is too old for checkboxes. In a "
            "notebook cell, run: %pip install --upgrade xlsxwriter   "
            "(then restart the kernel)"
        )

    # Updated and Reprocessed come first, right before Employee Name.
    checkbox_columns = ["Updated", "Reprocessed"]
    offset = len(checkbox_columns)
    columns = checkbox_columns + columns
    updated_col = columns.index("Updated")
    reprocessed_col = columns.index("Reprocessed")
    last_row = len(rows)  # zero-based index of the last data row
    last_col = len(columns) - 1

    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})

    try:
        sheet = workbook.add_worksheet("Payroll Returns")

        base = {
            "bg_color": ROW_COLOR,
            "border": 1,
            "border_color": "#D9D9D9",
        }
        header_format = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": HEADER_COLOR,
        })
        # Text formats preserve leading zeros and treat values as literal text.
        text_format = workbook.add_format({**base, "num_format": "@"})
        amount_format = workbook.add_format({**base, "num_format": "#,##0.00"})
        date_format = workbook.add_format({**base, "num_format": "mm/dd/yyyy"})
        checkbox_format = workbook.add_format({
            **base,
            "align": "center",
            "valign": "vcenter",
        })
        done_format = workbook.add_format({"bg_color": DONE_COLOR})

        # Excel table (filter buttons, structured name), no banded rows.
        sheet.add_table(0, 0, last_row, last_col, {
            "name": "PayrollReturns",
            "style": None,
            "banded_rows": False,
            "columns": [
                {"header": name, "header_format": header_format}
                for name in columns
            ],
        })

        for row_index, row in enumerate(rows, start=1):
            for data_index, value in enumerate(row):
                col_index = data_index + offset
                name = columns[col_index]

                if value is None:
                    sheet.write_blank(row_index, col_index, None, text_format)
                elif name == "Amount":
                    sheet.write_number(
                        row_index, col_index, value, amount_format
                    )
                elif name in ("Effective Date", "Report Date"):
                    sheet.write_datetime(
                        row_index, col_index, value, date_format
                    )
                else:
                    sheet.write_string(
                        row_index, col_index, str(value), text_format
                    )

            # Unchecked checkboxes.
            sheet.insert_checkbox(
                row_index, updated_col, False, checkbox_format
            )
            sheet.insert_checkbox(
                row_index, reprocessed_col, False, checkbox_format
            )

        # Whole row turns light green once both boxes are checked.
        updated_letter = xlsxwriter.utility.xl_col_to_name(updated_col)
        reprocessed_letter = xlsxwriter.utility.xl_col_to_name(reprocessed_col)
        sheet.conditional_format(1, 0, last_row, last_col, {
            "type": "formula",
            "criteria": (
                f"=AND(${updated_letter}2=TRUE,"
                f"${reprocessed_letter}2=TRUE)"
            ),
            "format": done_format,
        })

        widths = {
            "Updated": 12,
            "Reprocessed": 14,
            "Employee Name": 28,
            "Email": 24,
            "Original Trace Number": 25,
            "Returned Trace Number": 25,
            "Receiving Bank ID": 23,
            "Receiving Account": 24,
            "Amount": 15,
            "Reason Code": 16,
            "Reason Code Description": 45,
            "Effective Date": 18,
            "Report Date": 18,
        }

        for col_index, name in enumerate(columns):
            sheet.set_column(col_index, col_index, widths[name])

        sheet.freeze_panes(1, 0)

    finally:
        workbook.close()

    return buffer.getvalue()


def convert_payroll_pdf(root):
    # Select the PDF.
    selected_file = filedialog.askopenfilename(
        parent=root,
        title="Select a payroll PDF",
        initialdir=str(Path.cwd()),
        filetypes=[("PDF files", "*.pdf")],
    )

    if not selected_file:
        print("Canceled. Nothing was saved.")
        return

    pdf_path = Path(selected_file)

    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    # Read every page so employee records can span pages.
    reader = PdfReader(str(pdf_path))
    page_texts = [page.extract_text() or "" for page in reader.pages]

    columns, report_date, rows, skipped = extract_rows(page_texts)
    workbook_bytes = build_workbook_bytes(columns, rows)

    # Select the destination folder.
    selected_folder = filedialog.askdirectory(
        parent=root,
        title="Choose where to save the Excel file",
        initialdir=str(pdf_path.parent),
        mustexist=True,
    )

    if not selected_folder:
        print("Canceled. Nothing was saved.")
        return

    # Slashes cannot be used in Windows filenames.
    output_path = Path(selected_folder) / (
        f"BYU_ACH_Returns ({report_date:%m-%d-%Y}).xlsx"
    )

    # Ask before replacing an existing file.
    replace_existing = False

    if output_path.exists():
        replace_existing = messagebox.askyesno(
            parent=root,
            title="File already exists",
            message=(
                "There is an existing file with that name in your "
                "selected folder. Do you wish to replace it?"
            ),
            detail=(
                f"File: {output_path.name}\n"
                f"Folder: {output_path.parent}"
            ),
        )

        if not replace_existing:
            # Find the first available numbered filename.
            number = 1

            while True:
                numbered_path = output_path.with_name(
                    f"{output_path.stem}({number}).xlsx"
                )

                if not numbered_path.exists():
                    break

                number += 1

            accept_rename = messagebox.askyesno(
                parent=root,
                title="Save with a different name",
                message=(
                    "To save the file, we will rename it as "
                    f"{numbered_path.name}. Is that ok?"
                ),
            )

            if not accept_rename:
                print("Canceled. Nothing was saved.")
                return

            output_path = numbered_path

    if replace_existing:
        output_path.write_bytes(workbook_bytes)
    else:
        # Prevent overwriting a file that appears after checking.
        created_new_file = False

        try:
            with output_path.open("xb") as output_file:
                created_new_file = True
                output_file.write(workbook_bytes)
        except FileExistsError:
            raise FileExistsError(
                "A file with the selected name appeared before "
                "saving. Nothing was overwritten. Run the "
                "conversion again."
            )
        except Exception:
            if created_new_file:
                output_path.unlink(missing_ok=True)
            raise

    # Print confirmation without a popup that pauses the notebook.
    print(f"Created {len(rows)} employee row(s).")
    if skipped:
        print(f"Skipped {skipped} record(s) with a $0.00 amount.")
    print(f"Excel file saved to: {output_path}")


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        convert_payroll_pdf(root)

    except PermissionError:
        message = (
            "The file could not be accessed or saved. Close the "
            "Excel file if it is open, and make sure you have "
            "permission to access the PDF and destination folder."
        )
        print(message)
        messagebox.showerror(
            parent=root,
            title="Permission denied",
            message=message,
        )

    except Exception as error:
        print(f"Conversion failed: {error}")
        messagebox.showerror(
            parent=root,
            title="Conversion failed",
            message=str(error),
        )

    finally:
        root.destroy()


if __name__ == "__main__":
    main()