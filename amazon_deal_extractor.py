import argparse
import html
import json
import re
from dataclasses import dataclass
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_PROFILE_DIR = BASE_DIR / "data" / "amazon_playwright_profile"


@dataclass
class ExtractState:
    status: str
    reason: str
    body_text: str
    html: str


def _try_click(page, selectors: list[str], timeout_ms: int = 2500) -> bool:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=timeout_ms)
            locator.click(timeout=timeout_ms)
            return True
        except Exception:
            continue
    return False


def _try_fill(page, selectors: list[str], value: str, timeout_ms: int = 2500) -> bool:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=timeout_ms)
            locator.fill(value, timeout=timeout_ms)
            return True
        except Exception:
            continue
    return False


def _clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _first_match(text: str, patterns: list[str], flags: int = re.IGNORECASE | re.DOTALL) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            value = _clean_text(match.group(1))
            if value:
                return value
    return None


def _normalize_price(value: str | None) -> str | None:
    if not value:
        return None
    match = re.search(r"([$€£]?\s*[\d,]+(?:\.\d{2})?)", value)
    if not match:
        return None
    amount = match.group(1).replace(" ", "")
    if amount and amount[0].isdigit():
        amount = f"${amount}"
    return amount


def _normalize_image_url(value: str | None) -> str | None:
    if not value:
        return None
    url = value.strip()
    if not url:
        return None
    url = url.replace("\\/", "/").replace("\\u0026", "&")
    if url.startswith("//"):
        url = f"https:{url}"
    if url.startswith("http://") or url.startswith("https://"):
        return url
    return None


def _extract_selected_variant_from_twister(page_html: str) -> str | None:
    dimension_keys = ["color_name", "style_name", "pattern_name", "model_name"]
    for match in re.finditer(r'<script[^>]+type="a-state"[^>]*>(.*?)</script>', page_html, re.IGNORECASE | re.DOTALL):
        raw = (match.group(1) or "").strip()
        if not raw or ("sortedDimValuesForAllDims" not in raw and "selectedVariationValues" not in raw):
            continue
        try:
            payload = json.loads(html.unescape(raw))
        except Exception:
            continue
        if not isinstance(payload, dict):
            continue

        dims = payload.get("sortedDimValuesForAllDims")
        if isinstance(dims, dict):
            for dimension_key in dimension_keys:
                items = dims.get(dimension_key)
                if not isinstance(items, list):
                    continue
                for item in items:
                    if not isinstance(item, dict):
                        continue
                    if str(item.get("dimensionValueState") or "").upper() != "SELECTED":
                        continue
                    value = _clean_text(str(item.get("dimensionValueDisplayText") or ""))
                    if value:
                        return value

        selected = payload.get("selectedVariationValues")
        values = payload.get("variationValues")
        if isinstance(selected, dict) and isinstance(values, dict):
            for dimension_key in dimension_keys:
                selected_index = selected.get(dimension_key)
                dimension_values = values.get(dimension_key)
                if isinstance(selected_index, int) and isinstance(dimension_values, list):
                    if 0 <= selected_index < len(dimension_values):
                        value = _clean_text(str(dimension_values[selected_index] or ""))
                        if value:
                            return value
    return None


def _extract_from_html(html: str, body_text: str) -> dict[str, str | None]:
    discount_badge_scope = _first_match(
        html,
        [
            r'(<div id="dealBadge_feature_div".*?</div>)',
        ],
    )
    discount_type = None
    if discount_badge_scope:
        discount_type = _first_match(
            discount_badge_scope,
            [
                r">\s*((?:Limited|Prime Exclusive|Prime Big Deal|Lightning|Deal of the Day)[^<]{0,40}?deal)\s*<",
                r'"content"\s*:\s*\{\s*"fragments"\s*:\s*\[\s*\{\s*"text"\s*:\s*"([^"]*deal)"',
            ],
        )
    if not discount_type:
        discount_type = _first_match(
            html,
            [
                r'"messaging"\s*:\s*\{\s*"content"\s*:\s*\{\s*"fragments"\s*:\s*\[\s*\{\s*"text"\s*:\s*"([^"]*deal)"',
                r'"label"\s*:\s*\{\s*"content"\s*:\s*\{\s*"fragments"\s*:\s*\[\s*\{\s*"text"\s*:\s*"(\d+%\s+off)"',
            ],
        )

    price_scope = _first_match(
        html,
        [
            r'(<div id="corePriceDisplay_desktop_feature_div".*?<div id="tp-inline-twister-dim-values-container".*?</div>)',
            r'(<div id="apex_desktop".*?<div id="tp-inline-twister-dim-values-container".*?</div>)',
            r'(<div id="desktop_buybox".*?<div id="tp-inline-twister-dim-values-container".*?</div>)',
            r'(<div id="corePrice_feature_div".*?<div id="tp-inline-twister-dim-values-container".*?</div>)',
            r'(<div id="corePriceDisplay_desktop_feature_div".*?<div id="amazonGlobal_feature_div".*?</div>)',
            r'(<div id="apex_desktop".*?<div id="amazonGlobal_feature_div".*?</div>)',
            r'(<div id="desktop_buybox".*?<div id="buybox".*?</div>)',
        ],
    ) or html

    discount_strength = _first_match(
        price_scope,
        [
            r'"percentageDisplayString"\s*:\s*"(-?\d+%)"',
            r'"savingsPercentage"\s*:\s*"(-?\d+%)"',
            r'reinventPriceSavingsPercentageMargin[^>]*>\s*(?:<span[^>]*>)?\s*(-\d+%)\s*<',
            r"\b(-\d+%)\b",
            r"\b(\d+%\s+off)\b",
        ],
    )

    discount_price = _normalize_price(
        _first_match(
            price_scope,
            [
                r'apex-pricetopay-accessibility-label[^>]*>\s*\$?([\d,]+\.\d{2})\s+with\s+\d+\s+percent\s+savings',
                r'items\[0\.base\]\[customerVisiblePrice\]\[displayString\][^>]*value="\$?([\d,]+\.\d{2})"',
                r'customerVisiblePrice\]\[displayString\]"\s+value="\$?([\d,]+\.\d{2})"',
                r'"priceToPay"[^{}]{0,400}?"moneyValueOrRange":"([\d.,]+)"',
                r'"priceToPay"[^{}]{0,400}?"displayString":"[^$]*\$?([\d,]+\.\d{2})"',
                r'apex-pricetopay-accessibility-label[^>]*>\s*\$?([\d,]+\.\d{2})',
                r'Limited time deal.*?\$?([\d,]+\.\d{2})',
            ],
        )
    )

    list_price = _normalize_price(
        _first_match(
            price_scope,
            [
                r'"label"\s*:\s*"List Price:?"\s*,.*?"displayString"\s*:\s*"([^"]+)"',
                r'"listPrice"[^{}]{0,300}?"displayString":"([^"]+)"',
                r"List Price:?\s*</span>\s*<span[^>]*>\s*([$\d,]+\.\d{2})",
            ],
        )
    )
    if not list_price:
        list_price = _normalize_price(
            _first_match(
                body_text,
                [r"List Price:?\s*([$\d,]+\.\d{2})"],
                flags=re.IGNORECASE,
            )
        )

    typical_price = _normalize_price(
        _first_match(
            price_scope,
            [
                r'"basisPrice"\s*:\s*\{.*?"displayString"\s*:\s*"([^"]+)"',
                r'"label"\s*:\s*"Typical price:?"\s*,.*?"displayString"\s*:\s*"([^"]+)"',
                r'"label"\s*:\s*"Was Price:?"\s*,.*?"displayString"\s*:\s*"([^"]+)"',
                r"Typical price:?\s*</span>\s*<span[^>]*>\s*([$\d,]+\.\d{2})",
                r"Was Price:?\s*</span>\s*<span[^>]*>\s*([$\d,]+\.\d{2})",
            ],
        )
    )
    if not typical_price:
        typical_price = _normalize_price(
            _first_match(
                body_text,
                [
                    r"Typical price:?\s*([$\d,]+\.\d{2})",
                    r"Was Price:?\s*([$\d,]+\.\d{2})",
                ],
                flags=re.IGNORECASE,
            )
        )

    regular_price = _normalize_price(
        _first_match(
            price_scope,
            [
                r'"label"\s*:\s*"Regular Price:?"\s*,.*?"displayString"\s*:\s*"([^"]+)"',
                r'"regularPrice"[^{}]{0,300}?"displayString":"([^"]+)"',
                r"Regular Price:?\s*</span>\s*<span[^>]*>\s*([$\d,]+\.\d{2})",
            ],
        )
    )
    if not regular_price:
        regular_price = _normalize_price(
            _first_match(
                body_text,
                [r"Regular Price:?\s*([$\d,]+\.\d{2})"],
                flags=re.IGNORECASE,
            )
        )

    prime_member_price = _normalize_price(
        _first_match(
            price_scope,
            [
                r'"label"\s*:\s*"Prime Member Price:?"\s*,.*?"displayString"\s*:\s*"([^"]+)"',
                r'"primeMemberPrice"[^{}]{0,300}?"displayString":"([^"]+)"',
                r"Prime Member Price:?\s*</span>\s*<span[^>]*>\s*([$\d,]+\.\d{2})",
            ],
        )
    )
    if not prime_member_price:
        prime_member_price = _normalize_price(
            _first_match(
                body_text,
                [r"Prime Member Price:?\s*([$\d,]+\.\d{2})"],
                flags=re.IGNORECASE,
            )
        )

    color = _extract_selected_variant_from_twister(html) or _first_match(
        html,
        [
            r'"variation_color_name"\s*:\s*"([^"]+)"',
            r'"variation_style_name"\s*:\s*"([^"]+)"',
            r'id="inline-twister-row-color_name".*?<span[^>]*class="selection"[^>]*>\s*([^<]+)\s*<',
            r'id="inline-twister-row-style_name".*?<span[^>]*class="selection"[^>]*>\s*([^<]+)\s*<',
            r'"dimensionValuesDisplayData"\s*:\s*\{[^{}]*"color_name"\s*:\s*"([^"]+)"',
            r'"dimensionValuesDisplayData"\s*:\s*\{[^{}]*"style_name"\s*:\s*"([^"]+)"',
        ],
    )

    image_url = _normalize_image_url(
        _first_match(
            html,
            [
                r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']',
                r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image["\']',
                r'"hiRes"\s*:\s*"([^"]+)"',
                r'"large"\s*:\s*"([^"]+)"',
                r'"mainUrl"\s*:\s*"([^"]+)"',
            ],
        )
    )

    return {
        "image_url": image_url,
        "color": color,
        "discount_type": discount_type,
        "discount_strength": discount_strength,
        "discount_price": discount_price,
        "list_price": list_price,
        "typical_price": typical_price,
        "regular_price": regular_price,
        "prime_member_price": prime_member_price,
    }


def _read_page_state(page) -> ExtractState:
    html = page.content()
    body_text = _clean_text(page.locator("body").inner_text())

    if "opfcaptcha.amazon.com" in html or "Continue shopping" in body_text:
        return ExtractState("blocked", "amazon presented a verification or anti-bot page", body_text, html)

    if "No featured offers available" in body_text:
        return ExtractState("no_offer", "current session has no featured offer for this product", body_text, html)

    if "cannot be shipped to your selected delivery location" in body_text:
        return ExtractState("no_offer", "current delivery location cannot receive this item", body_text, html)

    if (
        "Add to Cart" in body_text
        or "Typical price" in body_text
        or "List Price" in body_text
        or "dealBadge_feature_div" in html
    ):
        return ExtractState("ready", "product page is available for extraction", body_text, html)

    return ExtractState("unknown", "page loaded but no reliable offer markers were found", body_text, html)


def _try_set_us_zip(page, zip_code: str) -> bool:
    if not _try_click(
        page,
        [
            "#nav-global-location-popover-link",
            "#glow-ingress-block",
            "#glow-ingress-line1",
        ],
    ):
        return False

    page.wait_for_timeout(1500)
    filled = _try_fill(
        page,
        [
            "#GLUXZipUpdateInput",
            "input[aria-label='or enter a US zip code']",
            "input[data-action='GLUXPostalUpdateAction']",
        ],
        zip_code,
    )
    if not filled:
        return False

    if not _try_click(
        page,
        [
            "#GLUXZipUpdate .a-button-input",
            "#GLUXZipUpdate-announce",
            "span[data-action='GLUXPostalUpdateAction'] input",
        ],
        timeout_ms=4000,
    ):
        page.keyboard.press("Enter")

    page.wait_for_timeout(3000)
    _try_click(
        page,
        [
            "input[data-action='GLUXConfirmClose']",
            "#GLUXConfirmClose",
            ".a-popover-footer input.a-button-input",
        ],
        timeout_ms=3000,
    )
    page.wait_for_timeout(2000)
    return True


def _auto_prepare_offer_context(page, us_zip: str) -> None:
    state = _read_page_state(page)
    if state.status == "blocked":
        return
    if (
        "cannot be shipped to your selected delivery location" in state.body_text
        or "No featured offers available" in state.body_text
        or "Deliver to Taiwan" in state.body_text
    ):
        if _try_set_us_zip(page, us_zip):
            page.wait_for_load_state("domcontentloaded", timeout=30000)
            page.wait_for_timeout(4000)


def _manual_gate(state: ExtractState) -> None:
    print()
    print(f"Current state: {state.status} - {state.reason}")
    print("If the browser shows a verification page, click through it.")
    print("If the item is unavailable for your location, switch the delivery address to a US address and select the target variant.")
    input("After the page looks correct in the browser, press Enter here to extract... ")


def _build_url(asin: str | None, url: str | None) -> str:
    if url:
        return url.strip()
    if not asin:
        raise ValueError("either --url or --asin is required")
    return f"https://www.amazon.com/dp/{asin.strip()}?th=1"


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract Amazon deal fields from a real browser session.")
    parser.add_argument("--url", help="Full Amazon product URL.")
    parser.add_argument("--asin", help="ASIN to build a default amazon.com URL.")
    parser.add_argument("--profile-dir", default=str(DEFAULT_PROFILE_DIR), help="Persistent browser profile directory.")
    parser.add_argument("--timeout-ms", type=int, default=90000, help="Navigation timeout in milliseconds.")
    parser.add_argument("--headless", action="store_true", help="Run headless. Not recommended for Amazon.")
    parser.add_argument("--manual", action="store_true", help="Pause for manual verification/address selection before extracting.")
    parser.add_argument("--debug-html", help="Optional output path for the final rendered HTML.")
    parser.add_argument("--us-zip", default="10001", help="US ZIP code to auto-apply when the current location has no offer.")
    args = parser.parse_args()

    url = _build_url(args.asin, args.url)
    profile_dir = Path(args.profile_dir).resolve()
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            channel="msedge",
            headless=args.headless,
            locale="en-US",
            timezone_id="America/Los_Angeles",
            viewport={"width": 1600, "height": 1400},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/148.0.0.0 Safari/537.36"
            ),
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=args.timeout_ms)
            page.wait_for_timeout(5000)
            _auto_prepare_offer_context(page, args.us_zip)
            state = _read_page_state(page)

            if args.manual:
                _manual_gate(state)
                try:
                    page.wait_for_timeout(1500)
                except PlaywrightTimeoutError:
                    pass
                state = _read_page_state(page)

            extracted = _extract_from_html(state.html, state.body_text)
            result = {
                "url": page.url,
                "status": state.status,
                "reason": state.reason,
                "discount_type": extracted["discount_type"],
                "discount_strength": extracted["discount_strength"],
                "discount_price": extracted["discount_price"],
                "list_price": extracted["list_price"],
                "typical_price": extracted["typical_price"],
                "regular_price": extracted["regular_price"],
                "prime_member_price": extracted["prime_member_price"],
            }

            if args.debug_html:
                Path(args.debug_html).write_text(state.html, encoding="utf-8")

            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0
        finally:
            context.close()


if __name__ == "__main__":
    raise SystemExit(main())
