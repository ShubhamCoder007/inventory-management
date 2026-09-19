"""Streamlit inventory manager for a small product catalog (Version 3).

stocks.xlsx is the single source of truth: the sheet is re-read on every interaction 
and written straight back after every mutation. The workbook keeps its own header 
names (product_name, cp, sp, ...); those are mapped onto the canonical names used 
below on read and mapped back on write.

Usage:
    streamlit run app.py
"""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd()
IMAGES_DIR = APP_DIR / "images"
DEFAULT_SHEET_NAME = "Sheet1"

# Whichever of these is on disk is the source of truth; the first name is used
# when a workbook has to be created from scratch.
WORKBOOK_NAMES = ("stocks.xlsx", "stock.xlsx")

# Canonical schema; this order is used when a fresh workbook is created.
COLUMNS = (
    "product_id",
    "name",
    "company",
    "type",
    "cost_price",
    "sell_price",
    "quantity",
    "image_path",
)
TABLE_COLUMNS = ["name", "company", "cost_price", "sell_price", "quantity"]

# Header spellings tolerated in a hand-made workbook, mapped to the canonical
# names above. Lookup happens after lowercasing and collapsing separators.
COLUMN_ALIASES = {
    "id": "product_id",
    "productid": "product_id",
    "sku": "product_id",
    "product": "name",
    "product_name": "name",
    "productname": "name",
    "item": "name",
    "item_name": "name",
    "category": "type",
    "product_type": "type",
    "brand": "company",
    "company_name": "company",
    "manufacturer": "company",
    "make": "company",
    "cp": "cost_price",
    "cost": "cost_price",
    "costprice": "cost_price",
    "buy_price": "cost_price",
    "purchase_price": "cost_price",
    "sp": "sell_price",
    "price": "sell_price",
    "sellprice": "sell_price",
    "sale_price": "sell_price",
    "selling_price": "sell_price",
    "qty": "quantity",
    "stock": "quantity",
    "stock_qty": "quantity",
    "image": "image_path",
    "imagepath": "image_path",
    "picture": "image_path",
    "photo": "image_path",
}

IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp", "bmp", "gif"]
NEW_TYPE_SENTINEL = "+ New type"
NEW_COMPANY_SENTINEL = "+ New company"

# Reorder levels live on a second sheet of the same workbook, one row per type.
THRESHOLD_SHEET = "thresholds"
THRESHOLD_COLUMNS = ("type", "min_quantity")
DEFAULT_THRESHOLD = 10

# st.dataframe's canvas renderer honours only background-color and color.
LOW_STOCK_STYLE = "background-color: #fdecea; color: #b3261e"


# ---------------------------------------------------------------------------
# Flash messages (survive the st.rerun() that follows every save)
# ---------------------------------------------------------------------------
def flash(kind: str, message: str) -> None:
    """Queue a message to be shown after the next rerun."""
    st.session_state.setdefault("flash", []).append((kind, message))


def render_flash() -> None:
    """Display and consume any pending flash messages."""
    for kind, message in st.session_state.pop("flash", []):
        getattr(st, kind)(message)


def apply_pending_selection() -> None:
    """Move queued selections onto their widget keys."""
    if "pending_type" in st.session_state:
        st.session_state["selected_type"] = st.session_state.pop("pending_type")
    if "pending_company" in st.session_state:
        st.session_state["selected_company"] = st.session_state.pop("pending_company")
    if "pending_product_id" in st.session_state:
        pid = st.session_state.pop("pending_product_id")
        st.session_state["selected_product_id"] = pid
        st.session_state["stock_product_id"] = pid


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------
def workbook_path() -> Path:
    """Return the existing workbook path, or default to the primary name."""
    for name in WORKBOOK_NAMES:
        if (APP_DIR / name).exists():
            return APP_DIR / name
    return APP_DIR / WORKBOOK_NAMES[0]


# Re-resolved on every rerun, so deleting the file mid-session is recoverable.
EXCEL_PATH = workbook_path()


def canonical_column(raw: object) -> str:
    """Map raw header string to canonical column name using COLUMN_ALIASES."""
    key = str(raw).strip().lower().replace(" ", "_").replace("-", "_")
    while "__" in key:
        key = key.replace("__", "_")
    return COLUMN_ALIASES.get(key, key)


def coerce_schema(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise headers and dtypes, adding any column the workbook lacks.

    Columns that are not part of the canonical schema are kept (and written back
    on save) so a workbook with extra bookkeeping columns is never truncated.
    """
    df = df.copy()
    df = df.rename(columns={col: canonical_column(col) for col in df.columns})
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(how="all")

    for column in COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA

    df["product_id"] = pd.to_numeric(df["product_id"], errors="coerce").astype("Int64")
    df["quantity"] = (
        pd.to_numeric(df["quantity"], errors="coerce").fillna(0).astype("Int64")
    )
    for column in ("cost_price", "sell_price"):
        df[column] = (
            pd.to_numeric(df[column], errors="coerce").fillna(0.0).astype(float)
        )
    for column in ("name", "company", "type", "image_path"):
        df[column] = (
            df[column]
            .astype("object")
            .where(df[column].notna(), "")
            .astype(str)
            .str.strip()
        )

    # Rows with no id still count as inventory, so give them one rather than
    # dropping them.
    unidentified = df["product_id"].isna()
    if unidentified.any():
        start = next_product_id(df)
        df.loc[unidentified, "product_id"] = range(start, start + int(unidentified.sum()))
        df["product_id"] = df["product_id"].astype("Int64")

    extras = [col for col in df.columns if col not in COLUMNS]
    return df[list(COLUMNS) + extras].reset_index(drop=True)


def blank_frame() -> pd.DataFrame:
    """Return an empty catalog DataFrame conforming to COLUMNS schema."""
    return coerce_schema(pd.DataFrame(columns=list(COLUMNS)))


def products_sheet(sheet_names: list[str]) -> str:
    """The catalog lives on the first sheet that is not the thresholds sheet."""
    for name in sheet_names:
        if name != THRESHOLD_SHEET:
            return name
    return DEFAULT_SHEET_NAME


def load_data() -> pd.DataFrame:
    """Read the workbook fresh, creating it with headers when absent or empty."""
    try:
        IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        st.error(f"Could not create the images folder: {exc}")

    try:
        missing_or_empty = not EXCEL_PATH.exists() or EXCEL_PATH.stat().st_size == 0
    except OSError as exc:
        st.error(f"Could not inspect `{EXCEL_PATH.name}`: {exc}")
        st.stop()

    if missing_or_empty:
        st.session_state["header_map"] = {}
        st.session_state["header_order"] = list(COLUMNS)
        fresh = blank_frame()
        if save_data(fresh):
            st.info(
                f"`{EXCEL_PATH.name}` was missing or empty, "
                "so a new workbook with just the headers was created."
            )
        return fresh

    try:
        with pd.ExcelFile(EXCEL_PATH, engine="openpyxl") as workbook:
            st.session_state["sheet_name"] = products_sheet(workbook.sheet_names)
            raw = workbook.parse(st.session_state["sheet_name"])
    except Exception as exc:  # unreadable / locked / corrupt workbook
        st.error(f"Could not read `{EXCEL_PATH.name}`: {exc}")
        st.stop()

    header_map: dict[str, str] = {}
    for header in raw.columns:
        header_map.setdefault(canonical_column(header), str(header))
    st.session_state["header_map"] = header_map
    st.session_state["header_order"] = [str(header) for header in raw.columns]

    return coerce_schema(raw)


def restore_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Put the workbook's own header spellings and column order back.

    The app works with canonical names internally, but the file on disk keeps
    whatever headers it shipped with (e.g. `cp` rather than `cost_price`) so
    anything else reading the workbook still sees the layout it expects.
    """
    original_names = st.session_state.get("header_map", {})
    original_order = st.session_state.get("header_order", [])
    renamed = df.rename(columns=original_names)
    ordered = [column for column in original_order if column in renamed.columns]
    ordered += [column for column in renamed.columns if column not in ordered]
    return renamed[ordered]


def load_thresholds() -> dict[str, int]:
    """Read the per-type reorder levels, or an empty mapping if unset."""
    try:
        if not EXCEL_PATH.exists() or EXCEL_PATH.stat().st_size == 0:
            return {}
        with pd.ExcelFile(EXCEL_PATH, engine="openpyxl") as workbook:
            if THRESHOLD_SHEET not in workbook.sheet_names:
                return {}
            sheet = workbook.parse(THRESHOLD_SHEET)
    except Exception as exc:
        st.error(f"Could not read the {THRESHOLD_SHEET} sheet: {exc}")
        return {}

    sheet = sheet.rename(columns={col: canonical_column(col) for col in sheet.columns})
    if not set(THRESHOLD_COLUMNS).issubset(sheet.columns):
        return {}

    levels = pd.to_numeric(sheet["min_quantity"], errors="coerce")
    return {
        str(name).strip(): max(0, int(level))
        for name, level in zip(sheet["type"], levels)
        if str(name).strip() and pd.notna(level)
    }


def save_data(df: pd.DataFrame, thresholds: dict[str, int] | None = None) -> bool:
    """Write both sheets back to Excel. Returns True when it stuck.

    Rewriting the file drops any sheet that is not re-written, so the catalog
    and the thresholds always go out together.
    """
    if thresholds is None:
        thresholds = st.session_state.get("thresholds", {})
    try:
        sheet_name = st.session_state.get("sheet_name", DEFAULT_SHEET_NAME)
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            restore_headers(df).to_excel(writer, sheet_name=sheet_name, index=False)
            if thresholds:
                pd.DataFrame(
                    {
                        "type": list(thresholds),
                        "min_quantity": list(thresholds.values()),
                    }
                ).to_excel(writer, sheet_name=THRESHOLD_SHEET, index=False)
        return True
    except Exception as exc:
        st.error(
            f"Could not save to `{EXCEL_PATH.name}`: {exc}\n\n"
            "If the workbook is open in Excel, close it and try again."
        )
        return False


def threshold_for(product_type: str, thresholds: dict[str, int]) -> int:
    """Return configured reorder threshold for a product type."""
    return thresholds.get(product_type, DEFAULT_THRESHOLD)


def threshold_series(df: pd.DataFrame, thresholds: dict[str, int]) -> pd.Series:
    """Per-row reorder level, resolved from each row's type."""
    return pd.Series(
        [threshold_for(value, thresholds) for value in df["type"]],
        index=df.index,
        dtype="int64",
    )


def next_product_id(df: pd.DataFrame) -> int:
    """Return next available unique integer product_id."""
    ids = pd.to_numeric(df.get("product_id"), errors="coerce").dropna()
    return int(ids.max()) + 1 if len(ids) else 1


# ---------------------------------------------------------------------------
# Pictures
# ---------------------------------------------------------------------------
def resolve_image(image_path: str) -> Path | None:
    """Return an existing file for a stored path, or None."""
    if not image_path:
        return None
    candidate = Path(image_path)
    if not candidate.is_absolute():
        candidate = APP_DIR / candidate
    try:
        return candidate if candidate.is_file() else None
    except OSError:
        return None


def save_image(uploaded, product_id: int) -> str | None:
    """Store an upload as images/<product_id>.<ext>; return its relative path."""
    suffix = Path(uploaded.name).suffix.lower() or ".png"
    target = IMAGES_DIR / f"{product_id}{suffix}"
    try:
        data = uploaded.getvalue()
        # Decode first so a broken upload fails before it reaches the workbook.
        with Image.open(io.BytesIO(data)) as probe:
            probe.verify()

        IMAGES_DIR.mkdir(parents=True, exist_ok=True)
        # A previous picture may sit under a different extension.
        for stale in IMAGES_DIR.glob(f"{product_id}.*"):
            if stale != target:
                stale.unlink(missing_ok=True)

        target.write_bytes(data)
        return target.relative_to(APP_DIR).as_posix()
    except Exception as exc:
        st.error(f"Could not save the picture: {exc}")
        return None


def is_new_upload(uploaded, state_key: str) -> bool:
    """True the first time a particular upload is seen.

    A file_uploader keeps its value across reruns, so without this guard the
    save-then-rerun cycle would re-save the same file forever.
    """
    token = getattr(uploaded, "file_id", None) or f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get(state_key) == token:
        return False
    st.session_state[state_key] = token
    return True


def placeholder_image(label: str) -> Image.Image:
    """Draw a simple 'no picture' tile so the layout does not jump around."""
    size = 420
    image = Image.new("RGB", (size, size), (240, 242, 246))
    draw = ImageDraw.Draw(image)
    draw.rectangle([6, 6, size - 7, size - 7], outline=(198, 203, 212), width=3)
    try:
        font = ImageFont.load_default(size=24)
    except TypeError:  # Pillow < 10.1 cannot size the default font
        font = ImageFont.load_default()
    for offset, text in ((-20, "No picture"), (20, label[:28])):
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        draw.text(
            ((size - (right - left)) / 2, (size - (bottom - top)) / 2 + offset),
            text,
            fill=(120, 127, 140),
            font=font,
        )
    return image


# ---------------------------------------------------------------------------
# Presentation helpers
# ---------------------------------------------------------------------------
def style_low_stock(display: pd.DataFrame, low: pd.Series):
    """Paint whole rows red where the quantity is under the reorder level."""
    styles = pd.DataFrame("", index=display.index, columns=display.columns)
    styles.loc[low, :] = LOW_STOCK_STYLE
    return display.style.apply(lambda _: styles, axis=None)


def product_labels(df: pd.DataFrame) -> dict[int, str]:
    """Name, company (when known) and id, so similar names stay tellable apart."""
    return {
        int(pid): " · ".join(
            part for part in (name or "(unnamed)", company, f"#{int(pid)}") if part
        )
        for pid, name, company in zip(df["product_id"], df["name"], df["company"])
    }


def choose_product(df: pd.DataFrame, state_key: str, label: str) -> int:
    """Render a product selectbox whose choice survives reruns.

    Each caller passes its own state_key, so the detail view and the stock
    update below it keep independent selections.
    """
    labels = product_labels(df)
    ids = list(labels)
    if st.session_state.get(state_key) not in ids:
        st.session_state[state_key] = ids[0]
    st.selectbox(label, ids, format_func=lambda pid: labels[pid], key=state_key)
    return int(st.session_state[state_key])


# ---------------------------------------------------------------------------
# Mutations
# ---------------------------------------------------------------------------
def commit(df: pd.DataFrame, message: str) -> None:
    """Persist the frame and rerun so every section sees the new numbers."""
    if save_data(df):
        flash("success", message)
        st.rerun()


def set_image_path(df: pd.DataFrame, product_id: int, relative_path: str) -> None:
    """Update relative image path for product ID and commit."""
    df.loc[df["product_id"] == product_id, "image_path"] = relative_path
    commit(df, "Picture updated.")


def change_quantity(df: pd.DataFrame, product_id: int, delta: int, label: str) -> None:
    """Adjust inventory stock quantity for product ID."""
    row = df.loc[df["product_id"] == product_id]
    if row.empty:
        st.error("That product is no longer in the workbook.")
        return

    current = int(row["quantity"].iloc[0])
    new_quantity = current + delta
    if new_quantity < 0:
        st.error(f"Cannot remove {abs(delta)} unit(s): only {current} in stock.")
        return

    df.loc[df["product_id"] == product_id, "quantity"] = new_quantity
    commit(df, f"{label}: {row['name'].iloc[0]} is now at {new_quantity} unit(s).")


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Inventory Manager", page_icon="📦", layout="wide")
st.title("📦 Inventory Manager")
st.caption(f"Reading and writing {EXCEL_PATH.name} in {APP_DIR}")
render_flash()
apply_pending_selection()

data = load_data()
thresholds = load_thresholds()
st.session_state["thresholds"] = thresholds  # so every save writes them back

# --- 1. Type & Company filter ------------------------------------------------
st.subheader("1 · Filter")
known_types = sorted({value for value in data["type"].tolist() if value})
known_companies = sorted({value for value in data["company"].tolist() if value})

# The two filters are independent: one company can sell across several types,
# so narrowing by type must not restrict the company choices or vice versa.
type_options = ["All"] + known_types
company_options = ["All"] + known_companies
if st.session_state.get("selected_type") not in type_options:
    st.session_state["selected_type"] = "All"
if st.session_state.get("selected_company") not in company_options:
    st.session_state["selected_company"] = "All"

filter_type_col, filter_company_col = st.columns(2)
selected_type = filter_type_col.selectbox(
    "Product type", type_options, key="selected_type"
)
selected_company = filter_company_col.selectbox(
    "Company", company_options, key="selected_company"
)

visible = data
if selected_type != "All":
    visible = visible[visible["type"] == selected_type]
if selected_company != "All":
    visible = visible[visible["company"] == selected_company]

# --- 2. Product table --------------------------------------------------------
st.subheader("2 · Products")
if visible.empty:
    st.info("No products match this filter. Add one further down the page.")
else:
    limits = threshold_series(visible, thresholds)
    low_stock = visible["quantity"] < limits

    left, middle, right, far_right = st.columns(4)
    left.metric("Products", len(visible))
    middle.metric("Units in stock", int(visible["quantity"].sum()))
    right.metric(
        "Stock value (at cost)",
        f"{(visible['cost_price'] * visible['quantity']).sum():,.2f}",
    )
    far_right.metric("Below threshold", int(low_stock.sum()))

    display = visible[TABLE_COLUMNS].copy()
    display["threshold"] = limits
    st.dataframe(
        style_low_stock(display, low_stock),
        hide_index=True,
        column_config={
            "name": st.column_config.TextColumn("Name"),
            "company": st.column_config.TextColumn("Company"),
            "cost_price": st.column_config.NumberColumn("Cost price", format="%.2f"),
            "sell_price": st.column_config.NumberColumn("Sell price", format="%.2f"),
            "quantity": st.column_config.NumberColumn("Quantity", format="%d"),
            "threshold": st.column_config.NumberColumn("Reorder below", format="%d"),
        },
    )
    if low_stock.any():
        st.error(
            "Low stock: "
            + ", ".join(
                f"{name} ({int(qty)} left, reorder below {int(limit)})"
                for name, qty, limit in zip(
                    visible.loc[low_stock, "name"],
                    visible.loc[low_stock, "quantity"],
                    limits[low_stock],
                )
            )
        )
    else:
        st.caption("Every product here is at or above its reorder level.")

# --- 2b. Reorder thresholds per type -----------------------------------------
with st.expander("Reorder thresholds by type"):
    if not known_types:
        st.caption("Thresholds appear once the catalog has at least one type.")
    else:
        st.caption(
            f"A product is flagged red when its quantity drops below the level "
            f"set for its type. Types with no level set use {DEFAULT_THRESHOLD}."
        )
        with st.form("threshold_form"):
            entered: dict[str, int] = {}
            grid = st.columns(4)
            for index, product_type in enumerate(known_types):
                entered[product_type] = int(
                    grid[index % 4].number_input(
                        product_type,
                        min_value=0,
                        step=1,
                        value=threshold_for(product_type, thresholds),
                        key=f"threshold_{product_type}",
                    )
                )

            if st.form_submit_button("Save thresholds", type="primary"):
                # Keep levels for types that are not currently in the catalog.
                merged = {**thresholds, **entered}
                if save_data(data, merged):
                    st.session_state["thresholds"] = merged
                    flash("success", "Reorder thresholds saved.")
                    st.rerun()

# --- 3. Product detail -------------------------------------------------------
st.subheader("3 · Product detail")
if visible.empty:
    st.caption("Nothing to inspect yet.")
else:
    product_id = choose_product(visible, "selected_product_id", "Product")
    product = data.loc[data["product_id"] == product_id].iloc[0]

    if st.button("View more", type="primary"):
        st.session_state["show_detail"] = True

    if st.session_state.get("show_detail"):
        picture_column, facts_column = st.columns([1, 2])
        existing = resolve_image(product["image_path"])

        with picture_column:
            if existing is not None:
                st.image(str(existing), caption=product["name"], use_container_width=True)
                uploader_label = "Update picture"
            else:
                if product["image_path"]:
                    st.warning(
                        f"`{product['image_path']}` is recorded but the file is missing."
                    )
                st.image(placeholder_image(product["name"]), use_container_width=True)
                uploader_label = "Upload picture"

            upload = st.file_uploader(
                uploader_label,
                type=IMAGE_EXTENSIONS,
                key=f"picture_upload_{product_id}",
            )
            if upload is not None and is_new_upload(
                upload, f"picture_token_{product_id}"
            ):
                stored = save_image(upload, product_id)
                if stored:
                    set_image_path(data, product_id, stored)

        with facts_column:
            margin = float(product["sell_price"]) - float(product["cost_price"])
            st.markdown(f"### {product['name'] or '(unnamed)'}")
            st.write(
                {
                    "Product ID": product_id,
                    "Company": product["company"] or "—",
                    "Type": product["type"] or "—",
                    "Cost price": f"{float(product['cost_price']):,.2f}",
                    "Sell price": f"{float(product['sell_price']):,.2f}",
                    "Margin per unit": f"{margin:,.2f}",
                    "Quantity": int(product["quantity"]),
                    "Image path": product["image_path"] or "—",
                }
            )

# --- 4. Stock update ---------------------------------------------------------
st.subheader("4 · Update stock")
if visible.empty:
    st.caption("Nothing to update yet.")
else:
    stock_id = choose_product(visible, "stock_product_id", "Product to adjust")
    stock_product = data.loc[data["product_id"] == stock_id].iloc[0]
    stock_limit = threshold_for(stock_product["type"], thresholds)
    on_hand = int(stock_product["quantity"])

    st.caption(
        f"**{stock_product['name'] or '(unnamed)'}** — {on_hand} unit(s) in stock, "
        f"reorder below {stock_limit}."
    )
    if on_hand < stock_limit:
        st.warning(f"This product is below its threshold of {stock_limit}.")

    delta = st.number_input(
        "Quantity change", min_value=1, step=1, value=1, key="quantity_delta"
    )
    add_column, sell_column, _ = st.columns([1, 1, 4])
    if add_column.button("Add stock", use_container_width=True):
        change_quantity(data, stock_id, int(delta), "Stock added")
    if sell_column.button("Record sale", use_container_width=True):
        change_quantity(data, stock_id, -int(delta), "Sale recorded")

# --- 5. Add product ----------------------------------------------------------
st.subheader("5 · Add a product")
with st.form("add_product", clear_on_submit=True):
    name_input = st.text_input("Name *", key="new_name")

    type_column, new_type_column = st.columns(2)
    type_choice = type_column.selectbox(
        "Type *", known_types + [NEW_TYPE_SENTINEL], key="new_type_choice"
    )
    new_type_input = new_type_column.text_input(
        "New type name",
        key="new_type_name",
        help=f"Used only when Type is set to “{NEW_TYPE_SENTINEL}”.",
    )

    company_picker, new_company_column = st.columns(2)
    company_choice = company_picker.selectbox(
        "Company *", known_companies + [NEW_COMPANY_SENTINEL], key="new_company_choice"
    )
    new_company_input = new_company_column.text_input(
        "New company name",
        key="new_company_name",
        help=f"Used only when Company is set to “{NEW_COMPANY_SENTINEL}”.",
    )

    cost_column, sell_column, quantity_column = st.columns(3)
    cost_input = cost_column.number_input(
        "Cost price", min_value=0.0, step=0.5, key="new_cost"
    )
    sell_input = sell_column.number_input(
        "Sell price", min_value=0.0, step=0.5, key="new_sell"
    )
    quantity_input = quantity_column.number_input(
        "Initial quantity", min_value=0, step=1, value=0, key="new_quantity"
    )

    image_input = st.file_uploader(
        "Picture (optional)", type=IMAGE_EXTENSIONS, key="new_product_picture"
    )

    if st.form_submit_button("Add product", type="primary"):
        product_name = name_input.strip()
        product_type = (
            new_type_input.strip()
            if type_choice == NEW_TYPE_SENTINEL
            else str(type_choice).strip()
        )
        company = (
            new_company_input.strip()
            if company_choice == NEW_COMPANY_SENTINEL
            else str(company_choice).strip()
        )

        problems = []
        if not product_name:
            problems.append("Name is required.")
        if not product_type:
            problems.append("Type is required — pick one or type a new one.")
        if not company:
            problems.append("Company is required — pick one or type a new one.")
        # Two companies may legitimately sell a product of the same name, so
        # only the company/name pair has to be unique.
        duplicate = data[
            (data["name"].str.casefold() == product_name.casefold())
            & (data["company"].str.casefold() == company.casefold())
        ]
        if product_name and company and not duplicate.empty:
            problems.append(f"{company} already has a product called “{product_name}”.")
        if float(sell_input) < 0:
            problems.append("Sell price must be zero or more.")
        if float(cost_input) < 0:
            problems.append("Cost price must be zero or more.")
        if int(quantity_input) < 0:
            problems.append("Initial quantity must be zero or more.")

        if problems:
            for problem in problems:
                st.error(problem)
        else:
            new_id = next_product_id(data)
            stored_path = save_image(image_input, new_id) if image_input else ""
            new_row = {
                "product_id": new_id,
                "name": product_name,
                "company": company,
                "type": product_type,
                "cost_price": float(cost_input),
                "sell_price": float(sell_input),
                "quantity": int(quantity_input),
                "image_path": stored_path or "",
            }
            updated = coerce_schema(
                pd.concat([data, pd.DataFrame([new_row])], ignore_index=True)
            )
            st.session_state["pending_type"] = product_type
            st.session_state["pending_company"] = company
            st.session_state["pending_product_id"] = new_id
            st.session_state["show_detail"] = True
            commit(updated, f"Added “{product_name}” as product #{new_id}.")
