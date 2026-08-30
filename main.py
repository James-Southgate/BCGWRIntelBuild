import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

import requests
import xlsxwriter
from dotenv import load_dotenv
from pypdf import PdfReader

from OneFlowSDK import OneflowSDK


OUTPUT_COLUMNS = [
    "Order",
    "Name / Certificate Text (or Fulfillment Item)",
    "RPC Certificate (Small)",
    "Record Holder (Large) Certificate",
    "Folder",
    "Medal",
    "Frame",
    "Pin Badge",
    "Additional Notes",
]

# These comparisons use both the item's code/SKU and its name/description.
# Add your exact Siteflow values here if they use different wording.
CERTIFICATE_TYPES = {
    "Record Holder (Large) Certificate": (
        "record holder",
        "large certificate",
        "large cert",
    ),
    "RPC Certificate (Small)": (
        "rpc",
        "small certificate",
        "small cert",
    ),
}

STOCK_TYPES = {
    "Folder": ("folder",),
    "Medal": ("medal",),
    "Frame": ("frame",),
    "Pin Badge": ("pin badge", "pin_badge", "pin-badge", "badge"),
}

SHIPMENT_COLOURS = [
    "FFF2CC",  # pale yellow
    "DDEBF7",  # pale blue
    "E2F0D9",  # pale green
    "FCE4D6",  # pale orange
    "E4DFEC",  # pale purple
    "DDEBF7",  # pale cyan/blue
]

ORDER_PAGE_SIZE = 100
MAX_ORDER_PAGES = 100


def application_directory():
    """Return the folder containing main.py or the bundled executable."""
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
        # Handle macOS .app bundle: Look 3 levels up to find the folder alongside the .app
        if sys.platform == "darwin" and app_dir.name == "MacOS" and app_dir.parent.name == "Contents":
            return app_dir.parent.parent.parent
        return app_dir
    return Path(__file__).resolve().parent


def parse_api_datetime(value):
    """Parse Siteflow ISO timestamps and normalise them to UTC."""
    if not value:
        return None

    text = str(value).strip()
    if text.endswith("Z"):
        text = f"{text[:-1]}+00:00"

    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    return parsed.astimezone(timezone.utc)


def order_datetime(order):
    """Find an order timestamp in either list or detail response shapes."""
    for key in ("date", "created", "createdAt", "orderTime"):
        parsed = parse_api_datetime(order.get(key))
        if parsed:
            return parsed

    nested = order.get("orderData")
    if isinstance(nested, dict):
        return order_datetime(nested)

    return None


def is_within_window(order_date, cutoff, now):
    return order_date is not None and cutoff <= order_date <= now


def load_config(base_directory):
    config_path = base_directory / "config.json"
    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}. "
            "Place config.json beside the executable."
        )

    with config_path.open("r", encoding="utf-8") as file:
        config = json.load(file)

    try:
        lookback_days = float(config.get("lookback_days", 1))
    except (TypeError, ValueError) as error:
        raise ValueError("config.json lookback_days must be a number") from error

    if lookback_days <= 0:
        raise ValueError("config.json lookback_days must be greater than zero")

    return {"lookback_days": lookback_days}


def output_path(base_directory, run_time):
    # Include seconds so two runs on the same day cannot overwrite one another.
    timestamp = run_time.strftime("%Y-%m-%d_%H%M%S")
    return base_directory / f"fulfilment_{timestamp}.xlsx"


def extract_order_page(payload):
    """Extract orders from common Siteflow list-response shapes."""
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return []

    for key in ("data", "orders", "items", "docs", "results"):
        value = payload.get(key)
        if isinstance(value, list):
            return value
        if isinstance(value, dict):
            for nested_key in ("data", "orders", "items", "docs", "results"):
                nested_value = value.get(nested_key)
                if isinstance(nested_value, list):
                    return nested_value
    return []


def find_total_pages(payload):
    """Read total-page metadata when the API supplies it."""
    if not isinstance(payload, dict):
        return None

    for key in ("totalPages", "pageCount", "pages"):
        value = payload.get(key)
        if isinstance(value, int):
            return value

    for key in ("paging", "pagination", "meta", "data"):
        nested = payload.get(key)
        value = find_total_pages(nested)
        if value is not None:
            return value
    return None


def find_total_items(payload):
    """Read total-record metadata when the API supplies it."""
    if not isinstance(payload, dict):
        return None

    for key in ("itemsTotal", "totalItems", "totalOrders", "totalRecords", "total"):
        value = payload.get(key)
        if isinstance(value, int):
            return value

    for key in ("paging", "pagination", "meta", "data"):
        nested = payload.get(key)
        value = find_total_items(nested)
        if value is not None:
            return value
    return None


def order_identity(order):
    for key in ("_id", "orderId", "id"):
        if order.get(key) is not None:
            return f"{key}:{order[key]}"
    return json.dumps(order, sort_keys=True, default=str)


def fetch_all_orders(client, cutoff):
    orders = []
    seen_order_ids = set()
    page = 1
    previous_order_date = None
    date_sort_is_valid = True

    for request_number in range(1, MAX_ORDER_PAGES + 1):
        params = {
            "page": page,
            "pagesize": ORDER_PAGE_SIZE,
        }
        response = client.request("GET", "/api/order", params=params)
        payload = json.loads(response.decode("utf-8"))
        page_orders = extract_order_page(payload)

        if not page_orders:
            break

        new_orders = []
        for order in page_orders:
            identity = order_identity(order)
            if identity not in seen_order_ids:
                seen_order_ids.add(identity)
                new_orders.append(order)

        if not new_orders:
            print(
                "WARNING: Siteflow returned the same order page twice. "
                "Pagination may not be behaving as expected."
            )
            break

        orders.extend(new_orders)
        print(
            f"Retrieved order page {page}: {len(page_orders)} records "
            f"({len(orders)} unique total)"
        )

        page_dates = [
            date for date in (order_datetime(order) for order in page_orders) if date
        ]
        for order_date in page_dates:
            if previous_order_date is not None and order_date > previous_order_date:
                date_sort_is_valid = False
            previous_order_date = order_date

        if not date_sort_is_valid:
            print(
                "WARNING: orders are not sorted newest-first; the program will "
                "continue paging instead of stopping at the date cutoff."
            )

        if (
            date_sort_is_valid
            and page_dates
            and page_dates[-1] < cutoff
        ):
            print(
                f"Reached orders older than the configured window on page {page}; "
                "no older pages are required."
            )
            break

        total_pages = find_total_pages(payload)
        total_items = find_total_items(payload)
        if total_pages is not None and request_number >= total_pages:
            break
        if total_items is not None and len(orders) >= total_items:
            break
        if (
            len(page_orders) < ORDER_PAGE_SIZE
            and (total_items is None or total_items <= len(orders))
        ):
            break

        page += 1
    else:
        print(
            f"WARNING: pagination stopped at max_pages="
            f"{MAX_ORDER_PAGES}."
        )

    return orders


def normalise(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def display_order_id(source_order_id, fallback):
    """Turn PW.34.1309.G.63460.5873007 into 63460."""
    source_order_id = str(source_order_id or "")

    match = re.search(r"(?:^|\.)G\.(\d+)(?:\.|$)", source_order_id, re.IGNORECASE)
    if match:
        return match.group(1)

    # Sensible fallback for similar dotted IDs without the explicit G marker.
    parts = source_order_id.split(".")
    if len(parts) >= 2 and parts[-2].isdigit():
        return parts[-2]

    return source_order_id or str(fallback)


def member_ids(member):
    return {
        str(value)
        for value in (
            member.get("_id"),
            member.get("itemId"),
            member.get("orderItemId"),
            member.get("stockId"),
            member.get("sourceItemId"),
        )
        if value is not None
    }


def shipment_id_for(member, shipments, member_type):
    """Resolve the shipment _id using references first and index second."""
    for key in ("shipmentId", "orderShipmentId"):
        if member.get(key):
            return str(member[key])

    ids = member_ids(member)
    reference_keys = (
        ("itemIds", "items")
        if member_type == "print_product"
        else ("stockItemIds", "stockItems")
    )

    for shipment in shipments:
        shipment_id = shipment.get("_id") or shipment.get("orderShipmentId")

        for key in reference_keys:
            for reference in shipment.get(key) or []:
                reference_ids = (
                    member_ids(reference)
                    if isinstance(reference, dict)
                    else {str(reference)}
                )
                if ids.intersection(reference_ids):
                    return str(shipment_id) if shipment_id else "unassigned"

    shipment_index = member.get("shipmentIndex")
    if shipment_index is not None:
        for list_index, shipment in enumerate(shipments):
            if shipment.get("shipmentIndex", list_index) == shipment_index:
                shipment_id = shipment.get("_id") or shipment.get("orderShipmentId")
                return str(shipment_id) if shipment_id else f"index-{shipment_index}"

    return "unassigned"


def classify(value, mappings):
    value = normalise(value)
    for column, aliases in mappings.items():
        if any(normalise(alias) in value for alias in aliases):
            return column
    return None


def certificate_column(item):
    searchable = " ".join(
        str(item.get(key) or "")
        for key in ("sku", "productDescription", "description")
    )
    return classify(searchable, CERTIFICATE_TYPES)


def stock_column(item):
    searchable = " ".join(
        str(item.get(key) or "") for key in ("code", "name")
    )
    return classify(searchable, STOCK_TYPES)


def artwork_paths(item):
    paths = []
    for component in item.get("components") or []:
        for key in ("path", "url", "filePath", "fetchUrl"):
            if component.get(key):
                paths.append(str(component[key]))
    return list(dict.fromkeys(paths))


def download_artwork(client, path):
    if path.lower().startswith(("http://", "https://")):
        response = requests.get(path, timeout=60)
        response.raise_for_status()
        return response.content

    api_path = path if path.startswith("/") else f"/{path}"
    return client.request("GET", api_path)


def extract_pdf_text(client, item):
    """Extract all selectable certificate text from the first readable PDF."""
    errors = []

    for path in artwork_paths(item):
        try:
            pdf_data = download_artwork(client, path)
            reader = PdfReader(BytesIO(pdf_data))
            pages = [page.extract_text() or "" for page in reader.pages]
            text = "\n".join(page.strip() for page in pages if page.strip()).strip()
            if text:
                return text, ""
            errors.append(f"No selectable text in {path}")
        except Exception as error:
            errors.append(f"Could not read {path}: {error}")

    return "", " | ".join(errors)


def blank_row(order_number):
    row = {column: "" for column in OUTPUT_COLUMNS}
    row["Order"] = order_number
    return row


def add_note(row, note):
    if not note:
        return
    existing = row["Additional Notes"]
    row["Additional Notes"] = f"{existing} | {note}" if existing else note


def add_quantity(row, column, quantity):
    current = row[column] or 0
    row[column] = current + (quantity or 0)


def build_rows(client, order_id, order_data):
    order_number = display_order_id(order_data.get("sourceOrderId"), order_id)
    shipments = order_data.get("shipments") or []
    grouped_certificates = defaultdict(list)

    # Each print item becomes a certificate row. Its PDF supplies the text.
    for item in order_data.get("items") or []:
        shipment_id = shipment_id_for(item, shipments, "print_product")
        row = blank_row(order_number)
        text, extraction_note = extract_pdf_text(client, item)
        row["Name / Certificate Text (or Fulfillment Item)"] = text

        column = certificate_column(item)
        if column:
            row[column] = item.get("quantity", 1) or 1
        else:
            description = item.get("productDescription") or item.get("description")
            add_note(
                row,
                f"Unmapped certificate type: {description or item.get('sku', '')}",
            )

        if extraction_note:
            add_note(row, extraction_note)

        grouped_certificates[shipment_id].append({"row": row, "item": item})

    # Stock items are folded into certificate rows; they never create new rows.
    for item in order_data.get("stockItems") or []:
        shipment_id = shipment_id_for(item, shipments, "stock_item")
        label = item.get("name") or item.get("code") or item.get("stockId", "Stock item")
        certificates = grouped_certificates.get(shipment_id) or []

        if not certificates:
            print(
                f"Order {order_number}: stock item {label!r} belongs to shipment "
                f"{shipment_id!r}, but that shipment has no certificate row."
            )
            continue

        # Prefer the certificate explicitly linked by sourceItemId. If Siteflow
        # only links the stock item to the shipment, place the shipment total on
        # its first certificate row rather than inventing a per-certificate split.
        source_item_id = item.get("sourceItemId")
        matching_certificates = [
            entry
            for entry in certificates
            if source_item_id is not None
            and entry["item"].get("sourceItemId") == source_item_id
        ]
        target = (matching_certificates or certificates)[0]["row"]

        column = stock_column(item)
        if column:
            add_quantity(target, column, item.get("quantity", 1) or 1)
        else:
            add_note(target, f"Unmapped stock item: {label}")

    # Make grouping visible in both colour and plain text.
    results = []
    for shipment_id, certificates in grouped_certificates.items():
        rows = [entry["row"] for entry in certificates]
        if len(rows) > 1:
            add_note(rows[0], "These pack/ship together")

        for row in rows:
            results.append((shipment_id, row))

    return results


def write_workbook(rows, filename):
    workbook = xlsxwriter.Workbook(str(filename))
    sheet = workbook.add_worksheet("Fulfilment")
    sheet.freeze_panes(1, 0)
    sheet.autofilter(0, 0, 0, len(OUTPUT_COLUMNS) - 1)

    header_format = workbook.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E78",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "border": 1,
    })
    sheet.write_row(0, 0, OUTPUT_COLUMNS, header_format)
    sheet.set_row(0, 32)

    row_formats = {}
    colour_by_group = {}
    next_colour = 0

    for row_number, (shipment_group, row) in enumerate(rows, start=1):
        group_key = (row["Order"], shipment_group)
        if group_key not in colour_by_group:
            colour_by_group[group_key] = SHIPMENT_COLOURS[next_colour % len(SHIPMENT_COLOURS)]
            next_colour += 1

        colour = colour_by_group[group_key]
        if colour not in row_formats:
            row_formats[colour] = workbook.add_format({
                "bg_color": f"#{colour}",
                "valign": "top",
                "text_wrap": True,
                "border": 1,
            })

        sheet.write_row(
            row_number,
            0,
            [row[column] for column in OUTPUT_COLUMNS],
            row_formats[colour],
        )
        sheet.set_row(row_number, 75)

    widths = [12, 65, 22, 30, 12, 12, 12, 14, 38]
    for column_index, width in enumerate(widths):
        sheet.set_column(column_index, column_index, width)

    workbook.close()


def main():
    base_directory = application_directory()
    load_dotenv(base_directory / ".env")
    config = load_config(base_directory)

    local_run_time = datetime.now().astimezone()
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(days=config["lookback_days"])

    client = OneflowSDK(
        os.environ["URL"],
        os.environ["TOKEN"],
        os.environ["SECRET"],
    )

    orders = fetch_all_orders(client, cutoff)

    rows = []
    included_orders = 0
    skipped_old = 0
    skipped_undated = 0

    for listed_order in orders:
        order_id = listed_order.get("_id")
        if not order_id:
            print(f"Skipping order with no usable ID: {listed_order}")
            continue

        # Filter before requesting the larger details response whenever the list
        # endpoint provides a date.
        listed_date = order_datetime(listed_order)
        if listed_date and not is_within_window(listed_date, cutoff, now):
            skipped_old += 1
            continue

        detail_path = f"/api/order/details/{quote(str(order_id), safe='')}"
        detail_response = client.request("GET", detail_path)
        detail = json.loads(detail_response.decode("utf-8"))
        order_data = detail.get("order", {}).get("orderData", {})

        if not order_data:
            print(f"Skipping order {order_id}: no orderData")
            continue

        detailed_date = order_datetime(order_data) or listed_date
        if detailed_date is None:
            skipped_undated += 1
            print(f"Skipping order {order_id}: no usable order date")
            continue
        if not is_within_window(detailed_date, cutoff, now):
            skipped_old += 1
            continue

        rows.extend(build_rows(client, order_id, order_data))
        included_orders += 1

    workbook_path = output_path(base_directory, local_run_time)
    write_workbook(rows, workbook_path)
    print(
        f"Wrote {len(rows)} fulfilment rows from {included_orders} orders "
        f"to {workbook_path}"
    )
    print(
        f"Window: last {config['lookback_days']:g} days from {now.isoformat()} UTC; "
        f"skipped {skipped_old} outside the window and {skipped_undated} undated."
    )


if __name__ == "__main__":
    main()
