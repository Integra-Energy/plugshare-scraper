#!/usr/bin/env python3
"""Collect PlugShare locations with unavailable chargers from a state-sized area.

The public PlugShare website loads regional data in a normal anonymous browser
session. This script uses that browser session and subdivides busy map regions so
that a state scan is not limited to the first page of results.
"""

from __future__ import annotations

import argparse
import csv
import re
import time
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

from playwright.sync_api import Page, sync_playwright

API_URL = "https://api.plugshare.com/v3/locations/region"
DEFAULT_STATUSES = ("UNDER_REPAIR", "OUTOFORDER", "UNAVAILABLE")
SERVER_RESULT_CAP = 250
MINIMUM_CELL_SIZE = 0.025

# Approximate bounding boxes are appropriate here because the API returns charger
# points. The CSV retains coordinates, making a stricter GIS boundary filter easy
# to add later if edge precision matters.
STATE_BOUNDS = {
    "PA": (39.7198, 42.2699, -80.5199, -74.6895),
}
REGION_BOUNDS = {
    # Continental North America: Canada, the contiguous United States, and Mexico.
    # Alaska, Hawaii, and island territories can be scanned separately if needed.
    "north-america": (14.0, 70.0, -141.0, -52.0),
}

US_STATE_CODES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC",
}
CANADIAN_PROVINCE_CODES = {
    "AB", "BC", "MB", "NB", "NL", "NS", "NT", "NU", "ON", "PE", "QC", "SK", "YT",
}
MEXICAN_STATE_CODES = {
    "AGS", "BCN", "BCS", "CAM", "CHH", "CHP", "CMX", "COA", "COL", "DGO",
    "GRO", "GTO", "HGO", "JAL", "MEX", "MIC", "MOR", "NAY", "NLE", "OAX",
    "PUE", "QRO", "ROO", "SIN", "SLP", "SON", "TAB", "TAM", "TLA", "VER",
    "YUC", "ZAC",
}
SUBDIVISION_CODES = US_STATE_CODES | CANADIAN_PROVINCE_CODES | MEXICAN_STATE_CODES
CANADIAN_POSTAL_PROVINCES = {
    "A": "NL", "B": "NS", "C": "PE", "E": "NB", "G": "QC", "H": "QC",
    "J": "QC", "K": "ON", "L": "ON", "M": "ON", "N": "ON", "P": "ON",
    "R": "MB", "S": "SK", "T": "AB", "V": "BC", "X": "NT", "Y": "YT",
}
US_ZIP_RANGES = (
    (350, 369, "AL"), (995, 999, "AK"), (850, 865, "AZ"), (716, 729, "AR"),
    (900, 961, "CA"), (800, 816, "CO"), (60, 69, "CT"), (197, 199, "DE"),
    (320, 349, "FL"), (300, 319, "GA"), (967, 968, "HI"), (832, 838, "ID"),
    (600, 629, "IL"), (460, 479, "IN"), (500, 528, "IA"), (660, 679, "KS"),
    (400, 427, "KY"), (700, 714, "LA"), (39, 49, "ME"), (206, 219, "MD"),
    (10, 27, "MA"), (480, 499, "MI"), (550, 567, "MN"), (386, 397, "MS"),
    (630, 658, "MO"), (590, 599, "MT"), (680, 693, "NE"), (889, 898, "NV"),
    (30, 38, "NH"), (70, 89, "NJ"), (870, 884, "NM"), (100, 149, "NY"),
    (270, 289, "NC"), (580, 588, "ND"), (430, 459, "OH"), (730, 749, "OK"),
    (970, 979, "OR"), (150, 196, "PA"), (28, 29, "RI"), (290, 299, "SC"),
    (570, 577, "SD"), (370, 385, "TN"), (750, 799, "TX"), (840, 847, "UT"),
    (50, 59, "VT"), (220, 246, "VA"), (980, 994, "WA"), (247, 268, "WV"),
    (530, 549, "WI"), (820, 831, "WY"), (200, 205, "DC"), (6, 9, "PR"),
)
SUBDIVISION_NAMES = {
    "ALBERTA": "AB", "BRITISH COLUMBIA": "BC", "MANITOBA": "MB",
    "NEW BRUNSWICK": "NB", "NEWFOUNDLAND": "NL", "NOVA SCOTIA": "NS",
    "ONTARIO": "ON", "PRINCE EDWARD ISLAND": "PE", "QUEBEC": "QC",
    "QUÉBEC": "QC", "SASKATCHEWAN": "SK", "PUERTO RICO": "PR",
}

ROW_FIELDS = (
    "location_id",
    "name",
    "address",
    "state_or_province",
    "latitude",
    "longitude",
    "location_under_repair",
    "majority_network_id",
    "majority_network_name",
    "station_id",
    "network_id",
    "network_name",
    "outlet_id",
    "connector",
    "status",
    "kilowatts",
    "url",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scan PlugShare's anonymous web map for unavailable chargers."
    )
    parser.add_argument(
        "--state",
        choices=sorted(STATE_BOUNDS),
        help="State bounding box to scan, such as PA.",
    )
    parser.add_argument(
        "--region",
        choices=sorted(REGION_BOUNDS),
        help="Larger region to scan, such as north-america.",
    )
    parser.add_argument(
        "--all-statuses",
        action="store_true",
        help="Write every charger outlet instead of unavailable chargers only.",
    )
    parser.add_argument(
        "--state-default",
        default="PA",
        choices=sorted(STATE_BOUNDS),
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("plugshare_unavailable_chargers.csv"),
        help="CSV output path (default: %(default)s).",
    )
    parser.add_argument(
        "--status",
        action="append",
        dest="statuses",
        help="Outlet status to include. Repeat to add values; defaults cover unavailable chargers.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.15,
        help="Seconds to wait between regional requests (default: %(default)s).",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Try a hidden browser. PlugShare currently rejects its regional requests in headless mode.",
    )
    return parser.parse_args()


def dismiss_registration_prompt(page: Page) -> None:
    close_dialog = page.locator(
        "button[aria-label='cancel']:visible, md-dialog [aria-label='close']:visible"
    )
    if close_dialog.count():
        close_dialog.first.click(force=True)
        page.wait_for_timeout(500)


def start_anonymous_map_session(page: Page) -> str:
    auth_headers: list[str] = []

    def remember_web_client_auth(request) -> None:
        if "/v3/locations/region?" in request.url:
            authorization = request.headers.get("authorization")
            if authorization:
                auth_headers.append(authorization)

    page.on("request", remember_web_client_auth)
    page.goto("https://www.plugshare.com/", wait_until="domcontentloaded", timeout=30_000)
    page.wait_for_timeout(6_000)
    dismiss_registration_prompt(page)

    deadline = time.monotonic() + 30
    while not auth_headers and time.monotonic() < deadline:
        page.wait_for_timeout(500)
    if not auth_headers:
        raise RuntimeError(
            "PlugShare did not start an anonymous regional map request. "
            "Run without --headless and confirm the map loads normally."
        )
    return auth_headers[-1]


def fetch_region(
    page: Page,
    authorization: str,
    center_latitude: float,
    center_longitude: float,
    span_latitude: float,
    span_longitude: float,
) -> list[dict[str, Any]]:
    query = urlencode(
        {
            "access": 1,
            "count": 500,
            "latitude": center_latitude,
            "longitude": center_longitude,
            "minimal": 0,
            "spanLat": span_latitude,
            "spanLng": span_longitude,
        }
    )
    url = f"{API_URL}?{query}"
    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            result = page.evaluate(
                """async ({url, authorization}) => {
                    const response = await fetch(url, {
                        headers: { authorization, accept: "application/json" }
                    });
                    if (!response.ok) throw new Error(`PlugShare HTTP ${response.status}`);
                    return await response.json();
                }""",
                {"url": url, "authorization": authorization},
            )
            break
        except Exception as exc:
            last_error = exc
            if attempt == 3:
                raise
            wait_seconds = attempt * 2
            print(f"Warning: regional fetch failed on attempt {attempt}; retrying in {wait_seconds}s")
            page.wait_for_timeout(wait_seconds * 1000)
    else:
        raise RuntimeError("PlugShare regional fetch failed.") from last_error
    if not isinstance(result, list):
        raise RuntimeError("PlugShare returned an unexpected regional response shape.")
    return result


def fetch_networks(page: Page, authorization: str) -> dict[int, str]:
    result = page.evaluate(
        """async ({authorization}) => {
            const response = await fetch("https://api.plugshare.com/v3/networks", {
                headers: { authorization, accept: "application/json" }
            });
            if (!response.ok) throw new Error(`PlugShare networks HTTP ${response.status}`);
            return await response.json();
        }""",
        {"authorization": authorization},
    )
    if not isinstance(result, list):
        raise RuntimeError("PlugShare returned an unexpected networks response shape.")
    networks: dict[int, str] = {}
    for network in result:
        network_id = network.get("id")
        name = network.get("name")
        if network_id is not None and name:
            networks[int(network_id)] = name
    return networks


def collect_locations(
    page: Page,
    authorization: str,
    bounds: tuple[float, float, float, float],
    delay: float,
) -> dict[int, dict[str, Any]]:
    locations: dict[int, dict[str, Any]] = {}
    request_count = 0

    def scan(cell: tuple[float, float, float, float]) -> None:
        nonlocal request_count
        min_latitude, max_latitude, min_longitude, max_longitude = cell
        span_latitude = max_latitude - min_latitude
        span_longitude = max_longitude - min_longitude
        request_count += 1
        regional_locations = fetch_region(
            page,
            authorization,
            (min_latitude + max_latitude) / 2,
            (min_longitude + max_longitude) / 2,
            span_latitude,
            span_longitude,
        )
        print(
            f"Region {request_count}: {len(regional_locations)} locations "
            f"({span_latitude:.2f} x {span_longitude:.2f} degrees)"
        )
        for location in regional_locations:
            locations[location["id"]] = location

        if len(regional_locations) < SERVER_RESULT_CAP:
            time.sleep(delay)
            return
        if span_latitude <= MINIMUM_CELL_SIZE or span_longitude <= MINIMUM_CELL_SIZE:
            print("Warning: a minimum-size region still reached the result cap.")
            time.sleep(delay)
            return

        middle_latitude = (min_latitude + max_latitude) / 2
        middle_longitude = (min_longitude + max_longitude) / 2
        for child in (
            (min_latitude, middle_latitude, min_longitude, middle_longitude),
            (min_latitude, middle_latitude, middle_longitude, max_longitude),
            (middle_latitude, max_latitude, min_longitude, middle_longitude),
            (middle_latitude, max_latitude, middle_longitude, max_longitude),
        ):
            scan(child)

    scan(bounds)
    print(f"Collected {len(locations)} unique locations from {request_count} regional requests.")
    return locations


def matching_rows(
    locations: dict[int, dict[str, Any]],
    statuses: set[str],
    all_statuses: bool,
    networks: dict[int, str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for location in locations.values():
        matched_outlet = False
        for station in location.get("stations", []):
            for outlet in station.get("outlets", []):
                status = outlet.get("status")
                if not all_statuses and status not in statuses:
                    continue
                matched_outlet = True
                rows.append(make_row(location, station, outlet, networks))
        if not all_statuses and location.get("under_repair") and not matched_outlet:
            rows.append(make_row(location, {}, {"status": "UNDER_REPAIR"}, networks))
    return sorted(
        rows,
        key=lambda row: (
            row["state_or_province"],
            row["address"],
            row["name"],
            str(row["location_id"]),
            str(row["station_id"]),
            str(row["outlet_id"]),
        ),
    )


def state_or_province(address: str) -> str:
    # PlugShare addresses commonly use postal abbreviations immediately before a
    # ZIP or postal code. Fall back to any recognized code when postal data is absent.
    upper_address = address.upper()
    normalized_address = re.sub(r"[^A-Z0-9]+", " ", upper_address)
    for name, code in SUBDIVISION_NAMES.items():
        if name in upper_address:
            return code
    tokens = re.findall(r"\b[A-Z]{2,3}\b", normalized_address)
    for token in reversed(tokens):
        if token in SUBDIVISION_CODES:
            return token
    canadian_postal_code = re.search(r"\b([ABCEGHJKLMNPRSTVXY])\d[A-Z](?:\s?\d[A-Z]\d)?\b", normalized_address)
    if canadian_postal_code:
        return CANADIAN_POSTAL_PROVINCES[canadian_postal_code.group(1)]
    # Mexico and the United States both use five-digit postal codes. Avoid
    # interpreting addresses that explicitly identify Mexico as US ZIP codes.
    if not re.search(r"\bMEXICO\b|\bMÉXICO\b", upper_address):
        us_zip_code = re.search(r"\b(\d{5})(?:-\d{4})?\b", normalized_address)
        if us_zip_code:
            prefix = int(us_zip_code.group(1)[:3])
            for minimum, maximum, code in US_ZIP_RANGES:
                if minimum <= prefix <= maximum:
                    return code
    return ""


def network_name(networks: dict[int, str], network_id: Any) -> str:
    if network_id in ("", None):
        return ""
    try:
        return networks.get(int(network_id), "")
    except (TypeError, ValueError):
        return ""


def make_row(
    location: dict[str, Any],
    station: dict[str, Any],
    outlet: dict[str, Any],
    networks: dict[int, str],
) -> dict[str, Any]:
    station_network_id = station.get("network_id", "")
    majority_network_id = location.get("majority_network_id", "")
    return {
        "location_id": location.get("id", ""),
        "name": location.get("name", ""),
        "address": location.get("address", ""),
        "state_or_province": state_or_province(location.get("address", "")),
        "latitude": location.get("latitude", ""),
        "longitude": location.get("longitude", ""),
        "location_under_repair": location.get("under_repair", False),
        "majority_network_id": majority_network_id,
        "majority_network_name": network_name(networks, majority_network_id),
        "station_id": station.get("id", ""),
        "network_id": station_network_id,
        "network_name": network_name(networks, station_network_id),
        "outlet_id": outlet.get("id", ""),
        "connector": outlet.get("connector", ""),
        "status": outlet.get("status", ""),
        "kilowatts": outlet.get("kilowatts", ""),
        "url": location.get("url", ""),
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as output:
        writer = csv.DictWriter(output, fieldnames=ROW_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    args = parse_args()
    if args.state and args.region:
        raise SystemExit("Choose either --state or --region, not both.")
    bounds = REGION_BOUNDS[args.region] if args.region else STATE_BOUNDS[args.state or args.state_default]
    statuses = {status.upper() for status in (args.statuses or DEFAULT_STATUSES)}
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=args.headless)
        page = browser.new_page(viewport={"width": 1600, "height": 1000})
        authorization = start_anonymous_map_session(page)
        networks = fetch_networks(page, authorization)
        print(f"Loaded {len(networks)} PlugShare networks.")
        locations = collect_locations(page, authorization, bounds, args.delay)
        rows = matching_rows(locations, statuses, args.all_statuses, networks)
        write_csv(args.output, rows)
        browser.close()
    print(f"Wrote {len(rows)} matching charger rows to {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
