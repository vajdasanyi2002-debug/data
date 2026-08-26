import os
import sys
import time
from pathlib import Path

import openpyxl
import requests


EXCEL_PATH = Path("Date_produse_Stalco.xlsx")
WEBHOOK_URL = os.environ["ODOO_WEBHOOK_URL"]

BATCH_SIZE = 200
TIMEOUT = 120

# TESZT MÓD
TEST_BARCODE = "5901466110942"


def blank(value):
    return value is None or (
        isinstance(value, str) and not value.strip()
    )


def barcode_value(value):
    if blank(value):
        return None

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return None

    value = str(value).strip()

    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]

    return value or None


def number_value(value):
    if blank(value):
        return None

    if isinstance(value, str):
        value = value.strip().replace(",", ".")

    return float(value)


def text_value(value):
    if blank(value):
        return None

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    value = str(value).strip()

    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]

    return value or None


def load_items():
    wb = openpyxl.load_workbook(
        EXCEL_PATH,
        read_only=True,
        data_only=True,
    )

    ws = wb.active

    headers = [
        cell.value
        for cell in next(
            ws.iter_rows(
                min_row=1,
                max_row=1,
            )
        )
    ]

    required = [
        "Cod e bare/Cod furnizo",
        "Denumire Stalco engleza",
        "UM",
        "PA Euro",
        "weight",
        "l10n_ro_net_weight",
        "hs_code",
    ]

    for name in required:
        if name not in headers:
            raise RuntimeError(
                "Missing Excel column: %s" % name
            )

    columns = {
        name: headers.index(name)
        for name in headers
    }

    items = []
    seen = set()

    skipped_barcode = 0
    skipped_empty = 0

    for row in ws.iter_rows(
        min_row=2,
        values_only=True,
    ):
        barcode = barcode_value(
            row[
                columns[
                    "Cod e bare/Cod furnizo"
                ]
            ]
        )

        if not barcode:
            skipped_barcode += 1
            continue

        if barcode in seen:
            raise RuntimeError(
                "Duplicate barcode in Excel: %s"
                % barcode
            )

        seen.add(barcode)

        item = {
            "barcode": barcode,
            "supplier_code": barcode,
        }

        # BESZÁLLÍTÓI TERMÉKNÉV
        supplier_name = text_value(
            row[
                columns[
                    "Denumire Stalco engleza"
                ]
            ]
        )

        if supplier_name is not None:
            item["supplier_name"] = supplier_name

        # BESZERZÉSI ÁR EUR
        supplier_price = number_value(
            row[
                columns[
                    "PA Euro"
                ]
            ]
        )

        if supplier_price is not None:
            item["supplier_price"] = supplier_price

        # BESZERZÉSI MÉRTÉKEGYSÉG
        um = text_value(
            row[
                columns[
                    "UM"
                ]
            ]
        )

        if um is not None:
            item["um"] = um

        # BRUTTÓ SÚLY
        weight = number_value(
            row[
                columns[
                    "weight"
                ]
            ]
        )

        if weight is not None:
            item["weight"] = weight

        # NETTÓ SÚLY
        net_weight = number_value(
            row[
                columns[
                    "l10n_ro_net_weight"
                ]
            ]
        )

        if net_weight is not None:
            item["l10n_ro_net_weight"] = net_weight

        # HS KÓD
        hs_code = text_value(
            row[
                columns[
                    "hs_code"
                ]
            ]
        )

        if hs_code is not None:
            item["hs_code"] = hs_code

        # Ha barcode + supplier_code-on kívül nincs adat
        if len(item) <= 2:
            skipped_empty += 1
            continue

        items.append(item)

    wb.close()

    print("Prepared:", len(items))
    print("Skipped without barcode:", skipped_barcode)
    print("Skipped without data:", skipped_empty)

    return items


def send_batch(
    batch,
    batch_number,
    total_batches,
):
    response = requests.post(
        WEBHOOK_URL,
        json={
            "source": "github_stalco_product_data",
            "items": batch,
        },
        timeout=TIMEOUT,
    )

    if not response.ok:
        raise RuntimeError(
            "Batch %s/%s failed: HTTP %s - %s"
            % (
                batch_number,
                total_batches,
                response.status_code,
                response.text[:1000],
            )
        )

    print(
        "Batch %s/%s OK - %s products"
        % (
            batch_number,
            total_batches,
            len(batch),
        )
    )


def main():
    items = load_items()

    # TESZT: CSAK A KIVÁLASZTOTT TERMÉK
    items = [
        item
        for item in items
        if item.get("barcode") == TEST_BARCODE
    ]

    if not items:
        raise RuntimeError(
            "Test product not found in Excel: %s"
            % TEST_BARCODE
        )

    print(
        "TEST MODE - barcode:",
        TEST_BARCODE,
    )

    print(
        "TEST PAYLOAD:",
        items[0],
    )

    total_batches = (
        len(items)
        + BATCH_SIZE
        - 1
    ) // BATCH_SIZE

    for start in range(
        0,
        len(items),
        BATCH_SIZE,
    ):
        batch = items[
            start:start + BATCH_SIZE
        ]

        batch_number = (
            start // BATCH_SIZE
        ) + 1

        send_batch(
            batch,
            batch_number,
            total_batches,
        )

        time.sleep(0.15)

    print(
        "DONE - %s rows sent to Odoo."
        % len(items)
    )


if __name__ == "__main__":
    try:
        main()

    except Exception as exc:
        print(
            "ERROR:",
            exc,
            file=sys.stderr,
        )

        sys.exit(1)
