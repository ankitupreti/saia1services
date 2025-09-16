#!/usr/bin/env python3
"""
rejection_challenge_full_improved.py

All-in-one script that:
- Loads bids from an Excel file (default: bids_input.xlsx)
- Loads fallback locators from an Excel file (default: locators of regection.xlsx) OR falls back to builtin locators
- For each bid: searches, finds the bid card, extracts department & technical status
- If disqualified: clicks the small "eye" icon to read disqualification reason (popup)
- Opens Representation/Challenge Rejection popup ("Click here to submit"), fills textarea, uploads a supporting PDF if provided in the input sheet, submits and verifies the response
- Uses robust fallback locator retrying and cropping screenshots to the bid card area

Run locally (example):
  pip install playwright pandas openpyxl
  playwright install
  python rejection_challenge_full_improved.py --input bids_input.xlsx --locators "locators of regection.xlsx" --manual-login

Notes:
- This file intentionally keeps locators flexible. Provide your own Excel with columns: name | selector (multiple rows for each named locator allowed).
- Selectors may be XPath (starting with //) or CSS/text selectors (Playwright supports many formats).

"""

import argparse
import asyncio
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ------------- Defaults & config -------------
DEFAULT_INPUT = "bids_input.xlsx"
DEFAULT_LOCATORS_XLSX = "locators of regection.xlsx"
OUTPUT_DIR = Path("representation_reports")
SNAPSHOT_INTERVAL = 100  # Optimized: increased from 50
SHORT_TIMEOUT_MS = 1000  # Optimized: reduced from 1500
MID_TIMEOUT_MS = 3000    # Optimized: reduced from 4000
LONG_TIMEOUT_MS = 10000  # Optimized: reduced from 20000
MAX_PAGE_RETRIES = 2
SLOW_MO = 0

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ------------- Helper: load locators from Excel or fallback -------------

def load_locators(path: Path) -> Dict[str, List[str]]:
    """Return a dict: {name: [selector1, selector2, ...]}"""
    fallback = {
        "username_input": ["#username", "input[name='username']", "input[type='text']"],
        "password_input": ["#password", "input[name='password']", "input[type='password']"],
        "login_button": ["button[type='submit']", "input[type='submit']", "button"],
        "participated_checkbox": ["//label[contains(.,'Bids/RAs Already Submitted/Participated')]/input", "//input[@type='checkbox'][contains(@aria-label,'Participated')]", "//input[@type='checkbox'][@value='participated']"],
        "search_input": ["//input[@type='search']", "css=input[placeholder*='Bid'], css=input[aria-label*='search'], input[name*=\"search\"]"],
        "bid_card": ["//div[contains(@class,'bid-card') or contains(@class,'panel') or contains(@class,'well') or contains(@class,'card')]"] ,
        "expand_button": ["//button[contains(text(),'View BID Results')]", "//button[contains(.,'View BID Results')]", "//a[contains(text(),'View BID Results')]", "//button[contains(@class,'view-details')]", "[class*='view-details']"],
        "department_name": ["//div[contains(.,'Department Name')]/following-sibling::div", "//div[contains(@class,'department')]"] ,
        "technical_status": ["//text()[contains(.,'Technical Status')]/parent::*", "//span[contains(text(),'Technical Status')]/following::node()[1]", "//div[contains(.,'Technical Status')]"] ,
        "technical_eye": ["//span[contains(text(),'Technical Status')]/following-sibling::a", "//img[contains(@src,'eye') or contains(@class,'eye') or contains(@alt,'eye')]"] ,
        "technical_popup_heading": ["//h2[contains(text(),'Reason for Technical Evaluation') or contains(.,'Reason for Technical Evaluation')]"] ,
        "technical_popup_comment": ["//div[contains(@class,'modal-content')]//table//tr[1]//td[4]", "//div[contains(@class,'modal-content')]//td[contains(.,'Comment')]/following::td[1]"] ,
        "representation_link": ["//a[contains(.,'Click here to submit') or contains(.,'Click here to submit') or contains(text(),'Click here to submit')]", "text=Click here to submit"],
        "representation_textarea": ["//textarea[contains(@placeholder,'Describe representation') or contains(@id,'representation')]", "textarea"],
        "upload_input": ["//input[@type='file']", "input[type='file']"],
        "submit_btn": ["//button[contains(text(),'Submit') or contains(.,'Submit') or //button[@type='submit']]", "text=Submit"],
        "modal_close": ["//button[contains(@class,'close') or contains(text(),'Close')]", "//a[contains(@class,'close')]"] ,
        "representation_history": ["//a[contains(.,'Clarification History') or contains(.,'Representation History') or contains(.,'History')]", "text=View"],
        "already_commented": ["//div[contains(.,'Already commented') or contains(.,'Already submitted')]"] ,
        "popup_ok": ["//button[contains(text(),'OK') or contains(text(),'Ok') or contains(text(),'ok')]"] ,
        "validation_message": ["//div[contains(@class,'alert') or contains(@class,'validation')]"]
    }

    if not path.exists():
        logging.info(f"Locators file not found at {path}. Using fallback locators built into script.")
        return fallback

    try:
        df = pd.read_excel(path, dtype=str).fillna("")
    except Exception as e:
        logging.warning(f"Failed to read locators from {path}: {e}. Using fallback.")
        return fallback

    locs: Dict[str, List[str]] = {}
    # Expect columns: name, selector (one selector per row). If multiple rows have same name, we aggregate.
    # If your Excel has a different format, this will still try to infer: first two cols as name+selector.
    cols = [str(c).strip().lower() for c in df.columns]
    if "name" in cols and "selector" in cols:
        for _, r in df.iterrows():
            name = str(r[df.columns[cols.index('name')]]).strip()
            sel = str(r[df.columns[cols.index('selector')]]).strip()
            if not name or not sel:
                continue
            locs.setdefault(name, []).append(sel)
    else:
        # Fallback: use first column as name and second as selector
        if len(df.columns) >= 2:
            for _, r in df.iterrows():
                name = str(r[df.columns[0]]).strip()
                sel = str(r[df.columns[1]]).strip()
                if not name or not sel:
                    continue
                locs.setdefault(name, []).append(sel)
        else:
            logging.warning("Locators Excel has unexpected format. Using built-in fallback locators.")
            return fallback

    # merge fallback keys not in file
    for k, v in fallback.items():
        if k not in locs:
            locs[k] = v
    logging.info(f"Loaded {len(locs)} locator groups from {path}")
    return locs


# ------------- Utilities for Playwright selectors -------------

def is_xpath(sel: str) -> bool:
    s = sel.strip()
    return s.startswith("//") or s.startswith("(//") or s.lower().startswith("xpath=")


def norm_locator_arg(sel: str) -> str:
    """Return a Playwright locator argument — prefix with xpath= for XPath selectors."""
    s = sel.strip()
    if s.lower().startswith("xpath="):
        return s
    if is_xpath(s):
        return f"xpath={s}"
    return s


async def try_get_text(page, selectors: List[str], timeout_ms: int = MID_TIMEOUT_MS) -> str:
    """Try each selector and return first non-empty text."""
    for sel in selectors:
        arg = norm_locator_arg(sel)
        try:
            loc = page.locator(arg)
            # If element present
            if await loc.count() == 0:
                # try JS query for CSS only
                if not is_xpath(sel):
                    try:
                        text = await page.evaluate(f'document.querySelector("{sel}")?.textContent || ""')
                        if text and text.strip():
                            return text.strip()
                    except Exception:
                        pass
                continue

            # prefer inner_text
            try:
                txt = await loc.inner_text(timeout=timeout_ms)
                if txt and txt.strip():
                    return txt.strip()
            except Exception:
                pass
            try:
                txt = await loc.text_content(timeout=timeout_ms)
                if txt and txt.strip():
                    return txt.strip()
            except Exception:
                pass
        except Exception:
            continue
    return ""


async def try_click(page, selectors: List[str], timeout_ms: int = MID_TIMEOUT_MS, locs: Dict[str, List[str]] = None) -> bool:
    """Try clicking each selector until one succeeds. Includes self-healing for popups."""
    for sel in selectors:
        arg = norm_locator_arg(sel)
        try:
            loc = page.locator(arg)
            if await loc.count() == 0:
                # For CSS selectors, try JS click
                if not is_xpath(sel):
                    try:
                        await page.evaluate(f'document.querySelector("{sel}")?.click && document.querySelector("{sel}").click()')
                        await page.wait_for_timeout(300)
                        logging.debug(f"JS-click invoked for {sel}")
                        return True
                    except Exception:
                        pass
                # XPath via JS is more involved -> skip
                continue

            await loc.wait_for(state="attached", timeout=timeout_ms)
            try:
                await loc.scroll_into_view_if_needed(timeout=timeout_ms)
            except Exception:
                pass

            try:
                await loc.click(timeout=timeout_ms)
            except Exception as e:
                # SELF-HEAL: check for obstruction error
                err_str = str(e).lower()
                if locs and ('obscures' in err_str or 'cover' in err_str or 'not visible' in err_str):
                    logging.warning(f"Click for '{sel}' failed, may be obscured. Attempting to self-heal...")
                    await ensure_popup_closed(page, locs)
                    await page.wait_for_timeout(300)
                    # RETRY click
                    await loc.click(timeout=timeout_ms)
                    logging.info(f"Self-heal click retry successful for: {sel}")
                else:
                    # If not obstruction error, fallback to original JS click logic
                    if not is_xpath(sel):
                        try:
                            await page.evaluate(f'document.querySelector("{sel}")?.click && document.querySelector("{sel}").click()')
                        except Exception:
                            raise e # re-raise outer exception
                    else:
                        raise e # re-raise outer exception

            await page.wait_for_timeout(250)
            logging.debug(f"Clicked selector: {sel}")
            return True
        except PlaywrightTimeoutError:
            logging.debug(f"Timeout clicking {sel}")
            continue
        except Exception as e:
            logging.debug(f"Error clicking {sel}: {e}")
            continue
    logging.warning(f"All click candidates failed: {selectors}")
    return False


async def crop_screenshot_of_locator(page, selector: str, out_path: Path):
    try:
        arg = norm_locator_arg(selector)
        loc = page.locator(arg).first
        if await loc.count() == 0:
            # fallback to full-page
            await page.screenshot(path=str(out_path), full_page=True)
            return
        await loc.scroll_into_view_if_needed()
        box = await loc.bounding_box()
        if not box:
            await page.screenshot(path=str(out_path), full_page=True)
            return
        clip = {"x": max(0, box["x"] - 8), "y": max(0, box["y"] - 8), "width": box["width"] + 16, "height": box["height"] + 16}
        await page.screenshot(path=str(out_path), clip=clip)
    except Exception as e:
        logging.warning(f"Cropped screenshot failed ({selector}): {e}")
        try:
            await page.screenshot(path=str(out_path), full_page=True)
        except Exception:
            pass


# ------------- Helper functions for popup handling -------------
async def ensure_popup_closed(page, locs: Dict[str, List[str]]):
    """Closes any visible popups or modals to ensure a clean state."""
    logging.debug("Checking for any open popups...")
    # Combine locators from locs file with common fallback selectors
    close_selectors = (
        locs.get('modal_close', []) +
        locs.get('popup_ok', []) +
        [
            "//button[contains(text(),'Close') or contains(@class,'close')]",
            "//button[contains(text(),'OK') or contains(text(),'Ok') or contains(text(),'ok')]",
            "//button[@aria-label='Close' or @aria-label='close']",
            "//button[normalize-space()='×']"  # Close 'x' button
        ]
    )

    # Use a set to avoid duplicate selectors
    unique_selectors = set(close_selectors)

    for sel in unique_selectors:
        try:
            # Find all matching buttons that are visible
            all_buttons = await page.locator(norm_locator_arg(sel)).all()
            for btn in all_buttons:
                if await btn.is_visible():
                    try:
                        await btn.click(timeout=SHORT_TIMEOUT_MS)
                        logging.info(f"Closed a popup using selector: {sel}")
                        # Wait a bit for the UI to settle after closing
                        await page.wait_for_timeout(500)
                    except Exception as e:
                        logging.debug(f"Could not click popup close button for selector {sel}: {e}")
        except Exception:
            # Ignore errors if a selector is invalid or element not found
            continue

async def handle_popup_safely(page, open_icon_locator: str, popup_body_locator: str, locs: Dict[str, List[str]], timeout_ms: int = 5000):
    """Wrap popup handling in a safe function"""
    await ensure_popup_closed(page, locs)
    try:
        # Click the icon to open popup
        arg = norm_locator_arg(open_icon_locator)
        await page.locator(arg).click()
        await page.wait_for_timeout(2000)

        # Verify popup is visible
        popup_arg = norm_locator_arg(popup_body_locator)
        popup_body = page.locator(popup_arg)
        await popup_body.wait_for(state="visible", timeout=timeout_ms)
        logging.debug("Popup opened successfully")

        # Close popup before continuing
        await ensure_popup_closed(page, locs)

    except PlaywrightTimeoutError:
        logging.warning(f"Popup did not appear for locator: {open_icon_locator}")
        await ensure_popup_closed(page, locs)

# ------------- Core: process a single bid -------------
async def process_single_bid(page, bid_no: str, row_idx: int, df_results: pd.DataFrame, locs: Dict[str, List[str]], options: dict):
    logging.info(f"[{row_idx+1}] Processing bid: {bid_no}")
    out_screens = OUTPUT_DIR / "screenshots"
    out_screens.mkdir(parents=True, exist_ok=True)

    try:
        # Proactively close any unexpected popups before starting
        await ensure_popup_closed(page, locs)

        # Ensure participated checkbox is checked (best-effort)
        try:
            checkbox_loc = page.locator(norm_locator_arg(locs.get('participated_checkbox', [''])[0]))
            if await checkbox_loc.count() > 0:
                if not await checkbox_loc.is_checked():
                    await checkbox_loc.check(force=True)
                    logging.info("Checked 'Bids/RAs Already Submitted/Participated' filter")
                    await asyncio.sleep(0.4)
        except Exception as e:
            logging.warning(f"Could not ensure participated filter: {e}")

        search_term = bid_no.split('/')[-1] if '/' in bid_no else bid_no
        # Try filling search box (if found)
        try:
            if await page.locator(norm_locator_arg(locs.get('search_input', [''])[0])).count() > 0:
                await try_click(page, locs.get('search_input', []), locs=locs)
                await page.fill(norm_locator_arg(locs.get('search_input', [''])[0]), search_term)
                await page.keyboard.press('Enter')
                await asyncio.sleep(0.4)
        except Exception:
            logging.debug('Search input not usable; continuing...')

        # Wait for bid card
        for _ in range(8):
            if await page.locator(norm_locator_arg(locs.get('bid_card', [''])[0])).count() > 0:
                break
            await asyncio.sleep(0.3)

        # Debug HTML Dump: Temporarily add after search (remove after fix)
        if await page.locator(norm_locator_arg(locs.get('bid_card', [''])[0])).count() > 0:
            html = await page.locator(norm_locator_arg(locs.get('bid_card', [''])[0])).first.inner_html(timeout=5000)
            logging.info(f"Bid {bid_no} card HTML snippet: {html[:1000]}...")

        # Expand bid details if needed
        await try_click(page, locs.get('expand_button', []), locs=locs)
        await asyncio.sleep(1)

        # Find technical status (try direct first)
        technical_status = await try_get_text(page, locs.get('technical_status', []))

        # If status blank, try expand button then re-check
        if not technical_status:
            await try_click(page, locs.get('expand_button', []), locs=locs)
            await asyncio.sleep(0.6)
            technical_status = await try_get_text(page, locs.get('technical_status', []))

        dept_name = await try_get_text(page, locs.get('department_name', []))
        df_results.at[row_idx, 'department_name'] = dept_name
        df_results.at[row_idx, 'technical_status'] = technical_status

        if not technical_status:
            df_results.at[row_idx, 'status'] = 'Skipped'
            df_results.at[row_idx, 'skip_reason'] = 'Blank technical status'
            # save cropped screenshot of card if possible
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_blank_{ts}.png")
            return

        # normalized check
        if 'qualified' in technical_status.lower() or technical_status.strip().lower().startswith('qualified'):
            df_results.at[row_idx, 'status'] = 'Skipped'
            df_results.at[row_idx, 'skip_reason'] = 'Qualified - no representation'
            return

        # Already commented
        already = await try_get_text(page, locs.get('already_commented', []))
        if already and 'comment' in already.lower():
            df_results.at[row_idx, 'status'] = 'Skipped'
            df_results.at[row_idx, 'skip_reason'] = 'Already commented'
            return

        # If disqualified - try to open the small eye icon to read reason
        disq_reason = ''
        if 'disqual' in technical_status.lower():
            # Use safe popup handling
            await handle_popup_safely(
                page,
                locs.get('technical_eye', [''])[0],
                locs.get('technical_popup_comment', [''])[0],
                locs
            )
            # Extract disqualification reason after popup handling
            disq_reason = await try_get_text(page, locs.get('technical_popup_comment', []))
        df_results.at[row_idx, 'disqualification_reason'] = disq_reason

        # Check representation history - if present, skip
        try:
            hist_text = await try_get_text(page, locs.get('representation_history', []))
            if hist_text and hist_text.strip():
                df_results.at[row_idx, 'representation_description'] = hist_text
                df_results.at[row_idx, 'representation_status'] = 'History'
                df_results.at[row_idx, 'status'] = 'Skipped'
                df_results.at[row_idx, 'skip_reason'] = 'Representation already submitted (history)'
                return
        except Exception:
            pass

        # Open representation popup with safe handling
        await ensure_popup_closed(page, locs)
        if not await try_click(page, locs.get('representation_link', []), locs=locs):
            df_results.at[row_idx, 'status'] = 'Skipped'
            df_results.at[row_idx, 'skip_reason'] = 'Representation link missing/click failed'
            await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_no_rep_link_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            return

        # Wait for popup to be visible
        try:
            modal_locator = page.locator("//div[contains(@class,'modal-dialog')]")
            await modal_locator.wait_for(state="visible", timeout=10000)
        except PlaywrightTimeoutError:
            df_results.at[row_idx, 'status'] = 'Skipped'
            df_results.at[row_idx, 'skip_reason'] = 'Representation popup did not appear'
            return

        await asyncio.sleep(0.6)

        # Fill textarea
        comment_text = ''
        if 'custom_comment' in df_results.columns and str(df_results.at[row_idx, 'custom_comment']).strip():
            comment_text = str(df_results.at[row_idx, 'custom_comment']).strip()
        else:
            comment_text = f"""Representation for Bid {bid_no}:
Documents submitted are as per ATC. Kindly re-evaluate and accept our bid."""

        filled = False
        for sel in locs.get('representation_textarea', []):
            try:
                arg = norm_locator_arg(sel)
                if await page.locator(arg).count() > 0:
                    await page.fill(arg, comment_text)
                    filled = True
                    break
            except Exception:
                continue
        if not filled:
            df_results.at[row_idx, 'status'] = 'Failed'
            df_results.at[row_idx, 'skip_reason'] = 'Representation textarea not found'
            await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_no_textarea_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            return

        # Optional upload
        if 'supporting_doc' in df_results.columns and str(df_results.at[row_idx, 'supporting_doc']).strip():
            path_str = str(df_results.at[row_idx, 'supporting_doc']).strip()
            p = Path(path_str)
            if not p.exists():
                alt = Path.cwd() / p
                if alt.exists():
                    p = alt
            if p.exists():
                uploaded = False
                for sel in locs.get('upload_input', []):
                    try:
                        arg = norm_locator_arg(sel)
                        if await page.locator(arg).count() > 0:
                            await page.set_input_files(arg, str(p))
                            uploaded = True
                            await asyncio.sleep(0.3)
                            break
                    except Exception:
                        continue
                if not uploaded:
                    logging.warning(f"Upload input not found for bid {bid_no}; file not attached: {p}")
            else:
                logging.warning(f"Supporting doc path not found for {bid_no}: {path_str}")

        # Submit
        submitted = await try_click(page, locs.get('submit_btn', []), locs=locs)
        await asyncio.sleep(0.6)

        # Handle OK confirmation popup safely
        await ensure_popup_closed(page, locs)

        # verify: either validation_message or representation_history now visible
        success = False
        val = await try_get_text(page, locs.get('validation_message', []))
        if val and val.strip():
            success = True
        else:
            rh = await try_get_text(page, locs.get('representation_history', []))
            if rh and rh.strip():
                success = True

        if success:
            df_results.at[row_idx, 'status'] = 'Success'
            df_results.at[row_idx, 'skip_reason'] = ''
            df_results.at[row_idx, 'updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_success_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            logging.info(f"Submitted representation for {bid_no}")
            return
        else:
            df_results.at[row_idx, 'status'] = 'Unknown'
            df_results.at[row_idx, 'skip_reason'] = 'No confirmation after submit - manual check required'
            await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_no_confirm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            logging.warning(f"No clear confirmation after submitting for {bid_no}")
            return

    except Exception as exc:
        df_results.at[row_idx, 'status'] = 'Failed'
        df_results.at[row_idx, 'skip_reason'] = f'Exception: {exc}'
        await crop_screenshot_of_locator(page, locs.get('bid_card', [''])[0], out_screens / f"{bid_no.replace('/','_')}_exception_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        logging.error(f"Exception processing {bid_no}: {exc}")
        traceback.print_exc()
        return


# ------------- Main orchestrator -------------
async def main():
    parser = argparse.ArgumentParser(description='GeM Rejection Challenge automation (robust with locator fallbacks).')
    parser.add_argument('--input', '-i', default=DEFAULT_INPUT, help='Input Excel file path (bids)')
    parser.add_argument('--sheet', '-s', default=None, help='Sheet name (optional)')
    parser.add_argument('--locators', default=DEFAULT_LOCATORS_XLSX, help='Excel with locators (name, selector)')
    parser.add_argument('--headless', '-H', default='false', help='Headless: true/false')
    parser.add_argument('--manual-login', action='store_true', help='Pause and allow manual login')
    parser.add_argument('--username', help='(Optional) username for auto-login')
    parser.add_argument('--password', help='(Optional) password for auto-login')
    parser.add_argument('--save-every', type=int, default=SNAPSHOT_INTERVAL, help='Save snapshot every N bids')
    parser.add_argument('--verbosity', '-v', action='count', default=1, help='Increase verbosity')
    args = parser.parse_args()

    level = logging.DEBUG if args.verbosity and args.verbosity > 1 else logging.INFO
    logging.basicConfig(level=level, format='%(asctime)s | %(levelname)s | %(message)s')

    headless = str(args.headless).lower() in ('1', 'true', 'yes', 'y')

    input_path = Path(args.input)
    if not input_path.exists():
        logging.error(f'Input path not found: {input_path}')
        sys.exit(1)

    # Load bids
    try:
        if args.sheet:
            df_input = pd.read_excel(input_path, sheet_name=args.sheet, dtype=str).fillna("")
            sheet_used = args.sheet
        else:
            # try some common sheet names
            df_input = None
            for s in ["Representation", "representation", "Sheet1", 0]:
                try:
                    df_input = pd.read_excel(input_path, sheet_name=s, dtype=str).fillna("")
                    sheet_used = s
                    break
                except Exception:
                    continue
            if df_input is None:
                logging.error('Could not read any sheet from input Excel.')
                sys.exit(1)
    except Exception as e:
        logging.error(f'Failed to read Excel: {e}')
        traceback.print_exc()
        sys.exit(1)

    df_input.columns = [str(c).strip() for c in df_input.columns]

    # find bid column
    bid_col = None
    for v in df_input.columns:
        if str(v).strip().lower() in ('bid_no', 'bid no', 'bidno', 'bid number', 'bid_number', 'bid') or 'bid' in str(v).lower():
            bid_col = v
            break
    if not bid_col:
        logging.error('Could not find bid column in input file. Ensure a column contains bid number.')
        sys.exit(1)

    # ensure result columns
    required = ['status', 'skip_reason', 'technical_status', 'disqualification_reason', 'department_name', 'representation_description', 'representation_status', 'updated_at']
    for c in required:
        if c not in df_input.columns:
            df_input[c] = ''
    if 'custom_comment' not in df_input.columns:
        df_input['custom_comment'] = ''
    if 'supporting_doc' not in df_input.columns:
        df_input['supporting_doc'] = ''

    # load locators
    locs = load_locators(Path(args.locators))

    total = len(df_input)
    logging.info(f'Starting processing {total} bids (sheet {sheet_used})')

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless, slow_mo=SLOW_MO, args=["--start-maximized"])
        context = await browser.new_context(viewport={"width": 1366, "height": 768})
        page = await context.new_page()

        await page.goto('https://bidplus.gem.gov.in/seller-bids', timeout=LONG_TIMEOUT_MS)
        logging.info('Opened seller-bids page')

        logging.info('Waiting for login page to be ready...')
        try:
            username_selectors = locs.get('username_input', [])
            if username_selectors:
                # Create a chain of locators with .or_() to find the first available username input
                locator_chain = page.locator(norm_locator_arg(username_selectors[0]))
                for sel in username_selectors[1:]:
                    locator_chain = locator_chain.or_(page.locator(norm_locator_arg(sel)))

                await locator_chain.first.wait_for(state="visible", timeout=LONG_TIMEOUT_MS)
                logging.info("Login page is ready (username input is visible).")
            else:
                # Fallback if no username locators are defined
                await page.wait_for_load_state('domcontentloaded', timeout=LONG_TIMEOUT_MS)
                logging.info("Login page is likely ready (DOM content loaded).")

        except PlaywrightTimeoutError:
            logging.warning("Timed out waiting for the login page to become ready. The script might fail at login.")
        except Exception as e:
            logging.error(f"An unexpected error occurred while waiting for the login page: {e}")

        if args.manual_login:
            print('\n>>> MANUAL LOGIN MODE: Pausing for 60 seconds to allow for manual login.')
            print('>>> Please complete login and navigate to the seller bids page.')
            await asyncio.sleep(60)
            print('\n>>> Resuming automation. If the page is ready, please press ENTER to continue...')
            input()
        elif args.username and args.password:
            try:
                if await page.locator(norm_locator_arg(locs.get('username_input', [''])[0])).count() > 0:
                    await page.fill(norm_locator_arg(locs.get('username_input', [''])[0]), args.username)
                if await page.locator(norm_locator_arg(locs.get('password_input', [''])[0])).count() > 0:
                    await page.fill(norm_locator_arg(locs.get('password_input', [''])[0]), args.password)
                await try_click(page, locs.get('login_button', []), locs=locs)
                await page.wait_for_load_state('networkidle', timeout=LONG_TIMEOUT_MS)
            except Exception:
                print('\n>>> Auto-login failed; please login manually and press ENTER to continue...')
                input()

        options = {'no_screenshots': False}

        for idx in range(total):
            bid_no = str(df_input.at[idx, bid_col]).strip()
            if not bid_no:
                df_input.at[idx, 'status'] = 'Skipped'
                df_input.at[idx, 'skip_reason'] = 'Empty bid number'
                continue

            for attempt in range(MAX_PAGE_RETRIES):
                try:
                    if 'seller-bids' not in page.url:
                        try:
                            await page.goto('https://bidplus.gem.gov.in/seller-bids', timeout=LONG_TIMEOUT_MS)
                        except Exception:
                            pass

                    await process_single_bid(page, bid_no, idx, df_input, locs, options)
                    break
                except Exception as e:
                    logging.error(f'Top-level exception for bid {bid_no} attempt {attempt+1}: {e}')
                    traceback.print_exc()
                    await page.screenshot(path=str(OUTPUT_DIR / 'screenshots' / f'top_exception_{bid_no.replace('/','_')}_{attempt+1}.png'), full_page=True)
                    await asyncio.sleep(0.6)
                    if attempt == MAX_PAGE_RETRIES - 1:
                        df_input.at[idx, 'status'] = 'Failed'
                        df_input.at[idx, 'skip_reason'] = f'Exception after {MAX_PAGE_RETRIES} attempts: {e}'

            if (idx + 1) % args.save_every == 0:
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                save_path = OUTPUT_DIR / f'snapshot_{ts}.xlsx'
                df_input.to_excel(save_path, index=False)

            await asyncio.sleep(0.2)

        # final save
        out_file = OUTPUT_DIR / f'processed_bids_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        df_input.to_excel(out_file, index=False)
        logging.info(f'All done. Saved: {out_file}')

        await context.close()
        await browser.close()


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logging.warning('Interrupted by user')
        sys.exit(1)
