#!/usr/bin/env python3
"""
VASN.py

Double-clickable Python script for Windows. The script reads Pipeline1.csv,
Pipeline2.csv, and Pipeline3.csv plus Received.csv from the current directory.
It normalizes Equipment IDs from the Pipeline files (removing the space and
leading zeros in the numeric part, e.g. "OMNX 002566" -> "OMNX2566"), then
compares them to the Received file's Railcar IDs and outputs a CSV containing
rows from all three Pipeline files for any railcars not present in Received.

Output file: Pipelines_missing_ASN.csv

No external dependencies (uses only Python standard library).
"""
import csv
import html
import os
import re
from pathlib import Path


def normalize_equipment_id(eid: str) -> str:
    """Normalize Equipment ID by removing spaces and stripping leading zeros
    from the final digit-group. Returns uppercase string.

    Examples:
    - "OMNX 002566" -> "OMNX2566"
    - "OMNX002566" -> "OMNX2566"
    - "AB12 000345" -> "AB120345"
    - If no trailing digits found, returns the string with whitespace removed and uppercased.
    """
    if eid is None:
        return ""
    s = eid.strip()
    if s == "":
        return ""
    # Remove all internal whitespace first
    s_nospace = re.sub(r"\s+", "", s)
    # Find last run of digits
    m = re.search(r"(\d+)$", s_nospace)
    if m:
        num = m.group(1).lstrip("0")
        if num == "":
            num = "0"
        prefix = s_nospace[: m.start(1)]
        return (prefix + num).upper()
    # If no trailing digits, just return the no-space uppercased value
    return s_nospace.upper()


def find_header_index(headers, candidates):
    """Return the header name from headers that matches any candidate (case-insensitive).
    If not found, returns None.
    """
    lower_map = {h.strip().lower(): h for h in headers}
    for c in candidates:
        key = c.strip().lower()
        if key in lower_map:
            return lower_map[key]
    return None


def read_received_ids(received_path):
    """Read the Received CSV and return a set of normalized Railcar IDs.
    Robustly finds the header row (in case of leading text) and normalizes each
    received ID using the same normalize_equipment_id function so comparisons are
    consistent.
    """
    ids = set()
    with open(received_path, newline="", encoding="utf-8-sig") as f:
        lines = [ln for ln in f]
    # find header row index by looking for a line that contains a Rail/Railcar header
    header_idx = None
    for i, ln in enumerate(lines):
        if 'rail' in ln.lower():
            header_idx = i
            break
    if header_idx is None:
        # fallback to first non-empty line
        for i, ln in enumerate(lines):
            if ln.strip():
                header_idx = i
                break
    if header_idx is None:
        return ids
    # Parse CSV from header row onward
    reader = csv.reader(lines[header_idx:])
    try:
        headers = next(reader)
    except StopIteration:
        return ids
    # find column index for rail id
    rc_col_name = find_header_index(headers, ["RailCar ID", "Railcar ID", "RailcarID", "Railcar"])
    if rc_col_name is not None:
        # get numeric index
        try:
            rc_idx = [h.strip().lower() for h in headers].index(rc_col_name.strip().lower())
        except ValueError:
            rc_idx = 0
    else:
        rc_idx = 0
    for row in reader:
        if len(row) <= rc_idx:
            continue
        raw = row[rc_idx]
        norm = normalize_equipment_id(raw)
        if norm:
            ids.add(norm)
    return ids


def process_pipeline(pipeline_path, received_ids, output_path=None):
    """Read Pipeline CSV (robust to leading title lines), normalize Equipment IDs,
    and write rows missing from received_ids to the output CSV. Returns count of
    missing rows and total rows processed.

    Output CSV will contain only these columns (if present):
      - Equipment ID (original)
      - _NormalizedEquipmentID
      - Shipper
      - Net Weight (lbs)
      - Gross Weight (lbs)
      - Tare Weight (lbs)
    """
    missing = []
    total = 0
    with open(pipeline_path, newline="", encoding="utf-8-sig") as f:
        lines = [ln for ln in f]
    # find header row index by looking for the Equipment ID header
    header_idx = None
    for i, ln in enumerate(lines):
        if 'equipment id' in ln.lower():
            header_idx = i
            break
    if header_idx is None:
        # fallback to first non-empty line
        for i, ln in enumerate(lines):
            if ln.strip():
                header_idx = i
                break
    if header_idx is None:
        return 0, 0
    reader = csv.reader(lines[header_idx:])
    try:
        headers = next(reader)
    except StopIteration:
        return 0, 0
    # normalize header names
    headers_clean = [h.strip() for h in headers]
    eq_col_name = find_header_index(headers_clean, ["Equipment ID", "EquipmentID", "Equipment"])
    if eq_col_name is not None:
        eq_idx = [h.lower() for h in headers_clean].index(eq_col_name.strip().lower())
    else:
        eq_idx = 0

    # helper to find header by substring (case-insensitive)
    def find_header_contains(headers_list, substrings):
        for h in headers_list:
            for s in substrings:
                if s.lower() in h.lower():
                    return h
        return None

    shipper_header = find_header_contains(headers_clean, ["Shipper", "Shipper Name"]) or None
    net_header = find_header_contains(headers_clean, ["Net Weight", "Net Weight (lbs)", "Ship Net"]) or None
    gross_header = find_header_contains(headers_clean, ["Gross Weight", "Gross Weight (lbs)", "Ship Gross"]) or None
    tare_header = find_header_contains(headers_clean, ["Tare Weight", "Tare Weight (lbs)", "Ship Tare"]) or None

    for row in reader:
        total += 1
        # ensure row is same length as headers
        if len(row) < len(headers_clean):
            # pad missing values with empty strings
            row = row + [''] * (len(headers_clean) - len(row))
        row_dict = {headers_clean[i]: row[i] for i in range(len(headers_clean))}
        raw_eid = row_dict.get(headers_clean[eq_idx], '')
        norm = normalize_equipment_id(raw_eid)
        if norm not in received_ids:
            # build compact output row (without normalized id)
            out_row = {}
            out_row[headers_clean[eq_idx]] = row_dict.get(headers_clean[eq_idx], '')
            # determine shipper value (empty string if not present)
            shipper_val = row_dict.get(shipper_header, '') if shipper_header else ''
            out_row[shipper_header if shipper_header else 'Shipper'] = shipper_val
            if net_header:
                out_row[net_header] = row_dict.get(net_header, '')
            if gross_header:
                out_row[gross_header] = row_dict.get(gross_header, '')
            if tare_header:
                out_row[tare_header] = row_dict.get(tare_header, '')
            missing.append(out_row)

    # Group / sort missing rows by shipper
    if missing:
        # Use shipper header key (or 'Shipper' placeholder)
        ship_key = shipper_header if shipper_header else 'Shipper'
        missing_sorted = sorted(missing, key=lambda x: (x.get(ship_key) or '').upper())
        # Build ordered out_headers
        out_headers = []
        out_headers.append(headers_clean[eq_idx])
        if ship_key not in out_headers:
            out_headers.append(ship_key)
        if net_header and net_header not in out_headers:
            out_headers.append(net_header)
        if gross_header and gross_header not in out_headers:
            out_headers.append(gross_header)
        if tare_header and tare_header not in out_headers:
            out_headers.append(tare_header)
        if output_path is not None:
            with open(output_path, "w", newline="", encoding="utf-8") as out_f:
                writer = csv.DictWriter(out_f, fieldnames=out_headers, extrasaction="ignore")
                writer.writeheader()
                for r in missing_sorted:
                    writer.writerow(r)
        # return missing rows and header info for follow-on processing
        return len(missing_sorted), total, missing_sorted, headers_clean[eq_idx], ship_key, net_header, gross_header, tare_header
    return len(missing), total, [], headers_clean[eq_idx] if headers_clean else headers_clean, (shipper_header if shipper_header else 'Shipper'), net_header, gross_header, tare_header


def process_pipelines(pipeline_paths, received_ids, output_path):
    """Process all pipeline files and write one deduplicated missing-ASN file."""
    all_missing = []
    total = 0
    seen_ids = set()
    metadata = None

    for pipeline_path in pipeline_paths:
        result = process_pipeline(pipeline_path, received_ids)
        _, file_total, missing_rows, equipment_header, shipper_key, net_header, gross_header, tare_header = result
        total += file_total
        if metadata is None:
            metadata = (equipment_header, shipper_key, net_header, gross_header, tare_header)
        for row in missing_rows:
            normalized_id = normalize_equipment_id(row.get(equipment_header, ""))
            if normalized_id and normalized_id not in seen_ids:
                seen_ids.add(normalized_id)
                all_missing.append(row)

    if metadata is None:
        return 0, total, [], "", "Shipper", None, None, None

    equipment_header, shipper_key, net_header, gross_header, tare_header = metadata
    all_missing.sort(key=lambda row: (row.get(shipper_key) or "").upper())
    if all_missing:
        out_headers = [equipment_header, shipper_key]
        for header in (net_header, gross_header, tare_header):
            if header and header not in out_headers:
                out_headers.append(header)
        with open(output_path, "w", newline="", encoding="utf-8") as out_f:
            writer = csv.DictWriter(out_f, fieldnames=out_headers, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(all_missing)

    return len(all_missing), total, all_missing, equipment_header, shipper_key, net_header, gross_header, tare_header


def verify_saved_outlook_draft(namespace, mail, drafts_folder, store):
    """Return diagnostics proving that Outlook persisted the saved draft."""
    entry_id = str(mail.EntryID)
    store_id = str(getattr(store, "StoreID", "") or "")
    if not entry_id:
        raise RuntimeError("Outlook returned an empty EntryID after Save()")

    verified = namespace.GetItemFromID(entry_id, store_id) if store_id else namespace.GetItemFromID(entry_id)
    verified_folder = str(getattr(getattr(verified, "Parent", None), "FolderPath", "") or "")
    expected_folder = str(getattr(drafts_folder, "FolderPath", "") or "")
    if expected_folder and verified_folder and verified_folder.lower() != expected_folder.lower():
        raise RuntimeError(
            f"saved item resolved to '{verified_folder}' instead of '{expected_folder}'"
        )

    return (
        f" verified EntryID={entry_id}; folder={verified_folder or expected_folder}; "
        f"subject={str(getattr(verified, 'Subject', '') or '')!r}; "
        f"to={str(getattr(verified, 'To', '') or '')!r}"
    )


def find_matching_outlook_draft(drafts_folder, recipients, cc_recipients, subject, body):
    """Find an existing matching draft to prevent duplicate Outlook items."""
    expected_recipients = [re.sub(r"\s+", "", address).lower() for address in recipients]
    expected_cc = [re.sub(r"\s+", "", address).lower() for address in cc_recipients]
    required_ids = {
        normalize_equipment_id(match.group(0))
        for match in re.finditer(r"\b[A-Za-z]{2,}\s*\d{3,}\b", body)
    }
    items = drafts_folder.Items
    for index in range(1, int(items.Count) + 1):
        try:
            item = items.Item(index)
            item_subject = str(getattr(item, "Subject", "") or "").strip()
            item_to = re.sub(r"\s+", "", str(getattr(item, "To", "") or "")).lower()
            item_cc = re.sub(r"\s+", "", str(getattr(item, "CC", "") or "")).lower()
            item_body = (
                str(getattr(item, "Body", "") or "")
                + " "
                + str(getattr(item, "HTMLBody", "") or "")
            )
            item_body_ids = {
                normalize_equipment_id(match.group(0))
                for match in re.finditer(r"\b[A-Za-z]{2,}\s*\d{3,}\b", item_body)
            }
            recipient_matches = all(
                address in item_to or address.split("@", 1)[0] in item_to
                for address in expected_recipients
            )
            cc_matches = all(
                address in item_cc or address.split("@", 1)[0] in item_cc
                for address in expected_cc
            )
            ids_match = required_ids.issubset(item_body_ids)
            if (
                item_subject == subject.strip()
                and recipient_matches
                and cc_matches
                and ids_match
            ):
                return item
        except Exception:
            continue
    return None


def select_store_drafts_folder(store, fallback_folder):
    """Prefer the mailbox's Outlook-defined Drafts folder over a name match."""
    try:
        default_folder = store.GetDefaultFolder(16)
        if default_folder is not None:
            return default_folder
    except Exception:
        pass
    return fallback_folder


def show_saved_draft_folder(outlook, drafts_folder):
    """Refresh the visible Outlook Explorer to the folder used for saving."""
    try:
        explorer = outlook.ActiveExplorer()
        if explorer is not None:
            explorer.CurrentFolder = drafts_folder
            explorer.Activate()
            return "; Outlook Explorer navigated to the saved Drafts folder"
        explorers = outlook.Explorers
        explorer = explorers.Add(drafts_folder, 0)
        explorer.Display()
        return "; Outlook opened a new Explorer on the saved Drafts folder"
    except Exception as error:
        return f"; WARNING: could not navigate Outlook Explorer ({type(error).__name__}: {error})"


def describe_saved_item(mail):
    """Return item metadata that remains useful when Outlook views are stale."""
    try:
        return (
            f" item_entry_id={str(getattr(mail, 'EntryID', '') or '')!r};"
            f" message_class={str(getattr(mail, 'MessageClass', '') or '')!r};"
            f" created={str(getattr(mail, 'CreationTime', '') or '')!r};"
            f" modified={str(getattr(mail, 'LastModificationTime', '') or '')!r}"
        )
    except Exception as error:
        return f" item metadata unavailable ({type(error).__name__}: {error})"


def describe_folder_identity(folder):
    """Return folder identifiers useful for comparing MAPI and Outlook views."""
    try:
        return (
            f" folder_name={str(getattr(folder, 'Name', '') or '')!r};"
            f" folder_path={str(getattr(folder, 'FolderPath', '') or '')!r};"
            f" entry_id={str(getattr(folder, 'EntryID', '') or '')!r}"
        )
    except Exception as error:
        return f" folder identity unavailable ({type(error).__name__}: {error})"


def load_outlook_signature(signature_hint=""):
    """Load the configured Outlook signature without opening a compose window."""
    try:
        import winreg

        signature_name = ""
        for version in ("16.0", "15.0", "14.0"):
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    f"Software\\Microsoft\\Office\\{version}\\Common\\MailSettings",
                ) as key:
                    signature_name = str(winreg.QueryValueEx(key, "NewSignature")[0] or "").strip()
                    if signature_name:
                        break
            except (FileNotFoundError, OSError):
                continue

        signature_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
        if not signature_name:
            candidates = list(signature_dir.glob("*.htm"))
            hint = re.sub(r"\s+", "", signature_hint).lower()
            matching = [
                path for path in candidates
                if hint and hint in re.sub(r"\s+", "", path.stem).lower()
            ]
            if len(matching) == 1:
                signature_path = matching[0]
            elif len(candidates) == 1:
                signature_path = candidates[0]
            else:
                print(
                    "DEBUG: Outlook signature name is not configured and "
                    f"{len(candidates)} signature files were found."
                )
                return ""
        else:
            signature_path = signature_dir / f"{signature_name}.htm"
        if not signature_path.exists():
            print(f"DEBUG: Outlook signature file not found: {signature_path}")
            return ""
        return signature_path.read_text(encoding="utf-8", errors="replace").strip()
    except Exception as error:
        print(f"DEBUG: Outlook signature could not be loaded: {type(error).__name__}: {error}")
        return ""


def connect_outlook_classic():
    """Connect only to the Outlook Classic COM automation server."""
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    version = str(getattr(outlook, "Version", "") or "")
    return outlook, namespace, version


def create_outlook_draft(
    namespace,
    drafts_folder,
    store,
    recipients,
    cc_recipients,
    subject,
    body,
    signature_hint="",
):
    """Create and save a draft directly in the requested Outlook folder."""
    step = "Create draft in target folder"
    try:
        existing = find_matching_outlook_draft(
            drafts_folder, recipients, cc_recipients, subject, body
        )
        if existing is not None:
            return existing, False, (
                f" Existing matching draft reused in "
                f"{getattr(drafts_folder, 'FolderPath', '<target Drafts folder>')}; "
                f"EntryID={str(getattr(existing, 'EntryID', '') or '')}"
            )
        mail = drafts_folder.Items.Add()
        step = "Resolve recipients"
        for address in recipients:
            recipient = mail.Recipients.Add(address)
            recipient.Type = 1
            if not recipient.Resolve():
                raise RuntimeError(f"Outlook could not resolve recipient '{address}'")
        for address in cc_recipients:
            recipient = mail.Recipients.Add(address)
            recipient.Type = 2
            if not recipient.Resolve():
                raise RuntimeError(f"Outlook could not resolve CC recipient '{address}'")
        if not mail.Recipients.ResolveAll():
            unresolved = [
                str(mail.Recipients.Item(i).Name)
                for i in range(1, mail.Recipients.Count + 1)
                if not mail.Recipients.Item(i).Resolved
            ]
            raise RuntimeError(f"Outlook could not resolve recipient(s): {', '.join(unresolved)}")
        step = "Set subject"
        mail.Subject = subject
        step = "Set message body"
        signature_html = load_outlook_signature(signature_hint)
        if signature_html:
            body_html = html.escape(body).replace("\n", "<br>")
            mail.HTMLBody = f"<html><body>{body_html}<br><br>{signature_html}</body></html>"
        else:
            mail.Body = body
        step = "Save message"
        mail.Save()
        verification = (
            f" Outlook created the item directly in "
            f"{getattr(drafts_folder, 'FolderPath', '<target Drafts folder>')};"
            + describe_saved_item(mail)
        )

        step = "Verify saved message"
        try:
            verification += ";" + verify_saved_outlook_draft(namespace, mail, drafts_folder, store)
        except Exception as verification_error:
            verification += (
                f"; WARNING: save completed, but post-save verification was unavailable "
                f"({type(verification_error).__name__}: {verification_error})"
            )
        return mail, True, verification
    except Exception as error:
        raise RuntimeError(f"{step} failed: {type(error).__name__}: {error}") from error


def generate_emails(missing_rows, equipment_header, shipper_key, contacts_path):
    """Generate Outlook drafts per shipper using ASNContacts.xlsx mapping.

    ASNContacts.xlsx expected columns:
      A: Friendly name
      B: Shipper code (matches shipper_key values in missing_rows)
      C: Email recipient(s) (comma/semicolon separated)
      D: Draft text (use *ASN* as placeholder for list of ASNs)
      E1: Shared CC recipient(s) (comma/semicolon/newline separated)
    """
    try:
        import openpyxl
    except ImportError:
        print("openpyxl is required to read ASNContacts.xlsx. Install with: pip install openpyxl")
        return
    if not os.path.exists(contacts_path):
        print(f"Contacts file not found: {contacts_path}")
        return
    try:
        wb = openpyxl.load_workbook(contacts_path, read_only=True)
    except Exception as e:
        print(f"Failed to open contacts file '{contacts_path}': {type(e).__name__}: {e}")
        return
    ws = wb.active
    contacts = {}
    cc_raw = ws["E1"].value
    cc_clean = re.sub(r"[;\n\r]+", ",", str(cc_raw or ""))
    cc_recipients = [address.strip() for address in cc_clean.split(",") if address.strip()]
    print(f"DEBUG: shared CC recipients={cc_recipients}")
    # Read all rows so headers can be detected and merged cells handled via fill-down
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return
    # Determine shipper codes present in missing_rows to help locate the code column
    present_shippers = set()
    for r in missing_rows:
        sv = (r.get(shipper_key) or '')
        present_shippers.add(re.sub(r"\s+", "", sv).upper())
    print(f"DEBUG: present_shippers={present_shippers}")

    # detect header row in first 3 rows (common labels)
    header_keywords = ('friendly', 'shipper', 'shipper code', 'shipper name', 'email', 'emails', 'draft')
    start_idx = 0
    for i in range(min(3, len(rows))):
        r = rows[i]
        if r:
            combined = ' '.join([str(c or '').lower() for c in r])
            if any(k in combined for k in header_keywords):
                start_idx = i + 1
                break

    # Build contacts by scanning rows and matching any cell to present_shippers.
    max_cols = max(len(r) for r in rows[start_idx:])
    # last seen values per column for fill-down behavior
    last_vals = [""] * max_cols
    for row in rows[start_idx:]:
        if not row:
            continue
        # update last_vals with current row's non-empty cells
        for col in range(len(row)):
            cell = row[col]
            if cell is not None and str(cell).strip() != "":
                last_vals[col] = str(cell).strip()
        # scan columns to see if any column now contains a shipper code we need
        for col in range(max_cols):
            val = last_vals[col]
            if not val:
                continue
            key = re.sub(r"\s+", "", val).upper()
            if key in present_shippers:
                # friendly assumed to be left of code, emails right of code, draft after emails
                friendly = last_vals[col-1] if col-1 >= 0 else ""
                emails = last_vals[col+1] if col+1 < max_cols else ""
                draft = last_vals[col+2] if col+2 < max_cols else ""
                contacts[key] = {"friendly": friendly or "", "emails": emails or "", "draft": draft or ""}
    print(f"DEBUG: contacts found for keys: {list(contacts.keys())}")
    # If contacts remains empty, nothing will be generated
    # group missing_rows by normalized shipper
    groups = {}
    norm_map = {}
    for r in missing_rows:
        ship_val = r.get(shipper_key, '')
        norm = re.sub(r"\s+", "", (ship_val or '')).upper()
        groups.setdefault(norm, []).append(r)
        norm_map[norm] = ship_val
    print(f"DEBUG: groups to generate: {list(groups.keys())}")
    # Try to create drafts in Outlook via COM (preferred)
    outlook_available = False
    outlook = None
    outlook_namespace = None
    try:
        outlook, outlook_namespace, outlook_version = connect_outlook_classic()
        outlook_available = True
        print(f"DEBUG: Outlook Classic COM available (version {outlook_version or 'unknown'})")
    except Exception as e:
        outlook_available = False
        print(
            "EMAIL FAILURE: Outlook Classic COM is unavailable. "
            "New Outlook does not support this automation method. "
            f"{type(e).__name__}: {e}"
        )
    drafts_count = 0

    # Use a hardcoded preferred store (mailbox) for Drafts per user request
    preferred_store = 'jmullins2@harsco.com'
    print(f"DEBUG: preferred_store hardcoded to: {preferred_store}")

    for norm_key, rows in groups.items():
        contact = contacts.get(norm_key)
        if not contact:
            print(f"No contact entry found for shipper '{norm_map.get(norm_key)}' (normalized '{norm_key}'), skipping email generation.")
            continue
        friendly = contact.get('friendly') or norm_map.get(norm_key)
        raw_addrs = (contact.get('emails') or '')
        # normalize separators: semicolons/newlines -> comma, then split
        addrs_clean = re.sub(r"[;\n\r]+", ",", raw_addrs)
        recipients = [a.strip() for a in addrs_clean.split(',') if a.strip()]
        if not recipients:
            print(f"EMAIL FAILURE [{friendly}]: no recipient addresses found for normalized shipper '{norm_key}'.")
            continue
        draft = contact.get('draft') or ''
        # collect ASNs (Equipment IDs)
        asns = [r.get(equipment_header, '').strip().strip('"') for r in rows if r.get(equipment_header)]
        asn_list_text = ', '.join(asns)
        # prepare body
        if '*ASN*' in draft:
            body = draft.replace('*ASN*', asn_list_text)
        else:
            body = (draft + "\n\nMissing ASNs: " + asn_list_text).strip()

        subject = f"Missing ASN(s) for {friendly}"

        if outlook_available:
            # Attempt to save into Drafts for every top-level store (user requested)
            stores_saved = []
            diagnostics = []
            ns = outlook_namespace
            try:
                preferred_store_attempted = False
                drafts_folder_found = False
                # If we detected a preferred_store from prior diagnostics, try that first
                if preferred_store:
                    diagnostics.append(f"Attempting preferred store: {preferred_store}")
                    # find store by name
                    target_store = None
                    for i in range(1, ns.Folders.Count + 1):
                        try:
                            s = ns.Folders.Item(i)
                            if str(s.Name).strip() == preferred_store:
                                target_store = s
                                break
                        except Exception:
                            continue
                    if target_store:
                        preferred_store_attempted = True
                        try:
                            # look for Drafts under target_store
                            drafts_folder = None
                            for fld in target_store.Folders:
                                try:
                                    if str(fld.Name).strip().lower() == 'drafts':
                                        drafts_folder = fld
                                        break
                                except Exception:
                                    continue
                            if drafts_folder:
                                drafts_folder = select_store_drafts_folder(target_store, drafts_folder)
                                drafts_folder_found = True
                                diagnostics.append(
                                    " Using preferred store Drafts folder:" + describe_folder_identity(drafts_folder)
                                )
                                try:
                                    count_before = int(drafts_folder.Items.Count)
                                except Exception:
                                    count_before = -1
                                diagnostics.append(f" Drafts folder path: {getattr(drafts_folder, 'FolderPath', str(drafts_folder))}; items before={count_before}")
                                _, created, verification = create_outlook_draft(
                                    ns,
                                    drafts_folder,
                                    target_store,
                                    recipients,
                                    cc_recipients,
                                    subject,
                                    body,
                                    preferred_store,
                                )
                                diagnostics.append(verification + show_saved_draft_folder(outlook, drafts_folder))
                                try:
                                    count_after = int(drafts_folder.Items.Count)
                                except Exception:
                                    count_after = -1
                                diagnostics.append(f" Saved draft in preferred store {preferred_store}; items after={count_after}")
                                if count_before >= 0 and count_after <= count_before:
                                    diagnostics.append(" WARNING: Outlook item count did not increase after Save(); verify the draft manually.")
                                stores_saved.append(preferred_store)
                                drafts_count += 1
                                action = "Saved" if created else "Reused"
                                print(
                                    f"{action} Outlook draft for {friendly} -> "
                                    f"To: {', '.join(recipients)}; CC: {', '.join(cc_recipients) or '<none>'} "
                                    f"(Drafts in '{preferred_store}')"
                                )
                            else:
                                diagnostics.append(f"Preferred store '{preferred_store}' has no Drafts folder")
                        except Exception as e:
                            diagnostics.append(
                                f"Failed while operating on preferred store '{preferred_store}': "
                                f"{type(e).__name__}: {e}"
                            )
                    else:
                        diagnostics.append(f"Preferred store '{preferred_store}' not found among Outlook stores")
                # If preferred_store was not successful (or not set), scan all stores
                # Do not retry a preferred-store operation in another store after
                # a COM failure: Save/Move may already have created the draft.
                if not stores_saved and not preferred_store_attempted:
                    for i in range(1, ns.Folders.Count + 1):
                        try:
                            store = ns.Folders.Item(i)
                            store_name = store.Name
                            diagnostics.append(f"Checking store: {store_name}")
                            # Try to find a Drafts folder under this store (case-insensitive search)
                            drafts_folder = None
                            for fld in store.Folders:
                                try:
                                    if str(fld.Name).strip().lower() == 'drafts':
                                        drafts_folder = fld
                                        break
                                except Exception:
                                    continue
                            # If not found, try one level deeper
                            if drafts_folder is None:
                                for fld in store.Folders:
                                    try:
                                        for sub in fld.Folders:
                                            try:
                                                if str(sub.Name).strip().lower() == 'drafts':
                                                    drafts_folder = sub
                                                    break
                                            except Exception:
                                                continue
                                        if drafts_folder is not None:
                                            break
                                    except Exception:
                                        continue
                            if not drafts_folder:
                                diagnostics.append(f" No Drafts folder found in store: {store_name}")
                                continue
                            try:
                                try:
                                    count_before = int(drafts_folder.Items.Count)
                                except Exception:
                                    count_before = -1
                                diagnostics.append(f" Drafts folder path: {getattr(drafts_folder, 'FolderPath', str(drafts_folder))}; items before={count_before}")
                                _, created, verification = create_outlook_draft(
                                    ns,
                                    drafts_folder,
                                    store,
                                    recipients,
                                    cc_recipients,
                                    subject,
                                    body,
                                    store_name,
                                )
                                diagnostics.append(verification + show_saved_draft_folder(outlook, drafts_folder))
                                try:
                                    count_after = int(drafts_folder.Items.Count)
                                except Exception:
                                    count_after = -1
                                diagnostics.append(f" Saved draft in store {store_name}; items after={count_after}")
                                if count_before >= 0 and count_after <= count_before:
                                    diagnostics.append(" WARNING: Outlook item count did not increase after Save(); verify the draft manually.")
                                stores_saved.append(store_name)
                                drafts_count += 1
                                action = "Saved" if created else "Reused"
                                print(
                                    f"{action} Outlook draft for {friendly} -> "
                                    f"To: {', '.join(recipients)}; CC: {', '.join(cc_recipients) or '<none>'} "
                                    f"(Drafts in '{store_name}')"
                                )
                            except Exception as inner_e:
                                diagnostics.append(f" Failed to save draft in store '{store_name}' for {friendly}: {inner_e}")
                        except Exception as store_e:
                            diagnostics.append(f" Error processing store index {i}: {store_e}")
                if not stores_saved and not drafts_folder_found:
                    diagnostics.append(f"No Outlook Drafts folder found for any store for {friendly}")
            except Exception as e:
                import traceback
                diagnostics.append(f"Outlook automation overall failure: {e}\n{traceback.format_exc()}")
            for diagnostic in diagnostics:
                print(f"EMAIL DIAGNOSTIC [{friendly}]: {diagnostic}")
        if not outlook_available:
            print(f"Outlook not available or failed; skipping draft creation for {friendly}.")
    # Summarize results
    print(f"Generated {drafts_count} Outlook draft(s).")


def main():
    # Use Pipeline1.csv, Pipeline2.csv, Pipeline3.csv, and Received.csv from the current directory.
    cwd = os.getcwd()
    pipeline_paths = [os.path.join(cwd, f"Pipeline{i}.csv") for i in range(1, 4)]
    received_path = os.path.join(cwd, "Received.csv")
    contacts_xlsx = os.path.join(cwd, "ASNContacts.xlsx")
    missing_pipeline_files = [path for path in pipeline_paths if not os.path.exists(path)]
    if missing_pipeline_files:
        for path in missing_pipeline_files:
            print(f"Pipeline file not found at: {path}")
        return
    if not os.path.exists(received_path):
        print(f"Received file not found at: {received_path}")
        return

    try:
        received_ids = read_received_ids(received_path)
    except Exception as e:
        print(f"Failed to read Received file: {e}")
        return

    base_dir = cwd
    output_name = "Pipelines_missing_ASN.csv"
    output_path = os.path.join(base_dir, output_name)

    try:
        missing_count, total, missing_rows, equipment_header, shipper_key, _, _, _ = process_pipelines(pipeline_paths, received_ids, output_path)
    except Exception as e:
        print(f"An error occurred while processing: {e}")
        return

    if missing_count == 0:
        print(f"Checked {total} pipeline rows. No missing ASNs were found.")
        # Remove output file if it was created but empty
        if os.path.exists(output_path) and os.path.getsize(output_path) == 0:
            try:
                os.remove(output_path)
            except Exception:
                pass
    else:
        msg = f"Found {missing_count} missing ASN(s) out of {total} pipeline rows. Saved results to: {output_path}"
        print(msg)
        # No automatic opening of folders; user requested no folder opening
        # If ASNContacts.xlsx exists, generate Outlook drafts grouped by shipper
        if os.path.exists(contacts_xlsx):
            generate_emails(missing_rows, equipment_header, shipper_key, contacts_xlsx)
        else:
            print(f"Contacts key not found at: {contacts_xlsx}. Skipping email draft generation.")


if __name__ == "__main__":
    main()
