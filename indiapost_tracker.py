"""
IndiaPost Tracking Script
Fetches the last tracking event for one or more postal article numbers.

Production base URL : https://app.indiapost.gov.in/beextcustomer/v1
Sandbox base URL    : https://test.cept.gov.in/beextcustomer/v1
Auth endpoint       : POST /access/login
Refresh endpoint    : POST /access/TokenWithRtoken
Tracking endpoint   : POST /tracking/bulk  (up to 50 articles)

Credentials are read from a .env file:
    INDIAPOST_ID=<your customer ID>
    INDIAPOST_SECRET=<your password>
"""

import sys
import os
import json
import time
import requests
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Load credentials from .env ────────────────────────────────────────────────
load_dotenv()
CLIENT_ID     = os.getenv("INDIAPOST_ID")
CLIENT_SECRET = os.getenv("INDIAPOST_SECRET")

# ── Configuration ─────────────────────────────────────────────────────────────
BASE_URL      = "https://app.indiapost.gov.in/beextcustomer/v1"
AUTH_URL      = f"{BASE_URL}/access/login"
REFRESH_URL   = f"{BASE_URL}/access/TokenWithRtoken"
TRACKING_URL  = f"{BASE_URL}/tracking/bulk"
# ─────────────────────────────────────────────────────────────────────────────


class AuthManager:
    """Handles login and automatic token refresh."""

    def __init__(self, username: str, password: str):
        self.username      = username
        self.password      = password
        self.access_token  = None
        self.refresh_token = None
        self.token_expiry  = 0

    def get_token(self) -> str:
        """Return a valid access token, refreshing or re-logging in as needed."""
        if not self.access_token or time.time() >= self.token_expiry - 60:
            if self.refresh_token:
                try:
                    self._refresh()
                    return self.access_token
                except Exception:
                    print("[INFO] Refresh failed — falling back to login...")
            else:
                print("[INFO] Token missing — logging in...")
            self.login()
        return self.access_token

    def _refresh(self):
        """Use the refresh token to get a new access token (no password needed)."""
        headers = {
            "Authorization": f"Bearer {self.refresh_token}",
            "Content-Type":  "application/json",
        }
        response = requests.post(REFRESH_URL, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()

        token_data        = data.get("data", data)  # some responses wrap in data
        new_token         = token_data.get("access_token")
        expires_in        = int(token_data.get("expires_in", 3600))

        if not new_token:
            raise RuntimeError(f"Token refresh returned no access_token: {data}")

        self.access_token = new_token
        self.token_expiry = time.time() + expires_in
        print(f"[OK] Token refreshed  (valid for {expires_in}s)")

    def login(self):
        """Authenticate with username/password and store tokens."""
        if not self.username or not self.password:
            raise ValueError(
                "Missing INDIAPOST_ID or INDIAPOST_SECRET in .env file."
            )

        payload = {"username": self.username, "password": self.password}
        headers = {"Content-Type": "application/json"}

        try:
            response = requests.post(
                AUTH_URL, json=payload, headers=headers, timeout=15
            )
            response.raise_for_status()
            data = response.json()

            if not data.get("success"):
                err = data.get("message") or data.get("error", {}).get("message", str(data))
                raise RuntimeError(f"Login failed: {err}")

            token_data         = data.get("data", {})
            self.access_token  = token_data.get("access_token")
            self.refresh_token = token_data.get("refresh_token")
            expires_in         = int(token_data.get("expires_in", 3600))
            self.token_expiry  = time.time() + expires_in

            print(f"[OK] Authenticated as {self.username}  "
                  f"(token valid for {expires_in}s)")

        except requests.exceptions.RequestException as e:
            print(f"[ERROR] Login request failed: {e}")
            if hasattr(e, "response") and e.response is not None:
                print(f"        Server response: {e.response.text[:300]}")
            raise


# ── Tracking helpers ──────────────────────────────────────────────────────────

def fetch_tracking(auth: AuthManager, article_numbers: list) -> dict:
    """Call the bulk tracking API and return the raw JSON response."""
    token   = auth.get_token()
    headers = {
        "Authorization":  f"Bearer {token}",
        "Content-Type":   "application/json",
    }
    payload  = {"bulk": article_numbers}
    response = requests.post(
        TRACKING_URL, json=payload, headers=headers, timeout=45
    )
    response.raise_for_status()
    return response.json()


def get_last_event(article_data: dict) -> dict | None:
    """Return the most recent tracking event (index 0 = newest), or None."""
    events = article_data.get("tracking_details", [])
    return events[0] if events else None


def determine_delivery_outcome(article: dict) -> str:
    """
    Classify the delivery result using the full event history.

    Statuses:
      Delivered to Addressee   — final delivery confirmed to intended recipient
      Delivered at Office      — some other delivered variant
      Returned to Sender       — item fully returned (final state)
      Return Journey           — delivery failed/attempted, now heading back
      Onward Journey           — in transit, no return events yet
      On Hold                  — item is on hold at a facility
      No Events / Not Found    — API returned nothing
    """
    events = article.get("tracking_details", [])
    if not events:
        return "No Events"

    latest_event = events[0].get("event", "")

    # ── Final delivery ────────────────────────────────────────────────────────
    if "Delivered(Addressee)" in latest_event:
        return "Delivered to Addressee"
    if "Delivered" in latest_event:
        return f"Delivered at Office"

    # ── Scan history for directional signals (skip pure logistics) ────────────
    LOGISTICS = {"Bag Close", "Bag Dispatch", "Bag Received", "Item Invoiced",
                 "Item Book", "Item Received"}
    directional = [
        e.get("event", "") for e in events
        if e.get("event", "") not in LOGISTICS
    ]

    if directional:
        top = directional[0]   # most recent meaningful event
        if "Onhold" in top or "On Hold" in top:
            return "On Hold"
        if "Returned to Sender" in top:
            return "Returned to Sender"
        if "Return" in top:
            return "Return Journey"

    return "Onward Journey"


def format_event(event: dict) -> str:
    """Format a single tracking event for display."""
    date_str = (event.get("date") or "")[:10] or "N/A"
    time_str = event.get("time",   "N/A")
    office   = event.get("office", "N/A")
    evt_name = event.get("event",  "N/A")
    return f"{date_str}  {time_str}  |  {evt_name}  |  {office}"


# ── Main function ─────────────────────────────────────────────────────────────

def track_articles(article_numbers: list, username: str = None, password: str = None):
    """Authenticate, call tracking API, and print the last event per article."""
    uid = username or CLIENT_ID
    pwd = password or CLIENT_SECRET

    if not uid or not pwd:
        print("[ERROR] No credentials found. "
              "Set INDIAPOST_ID and INDIAPOST_SECRET in your .env file.")
        return

    print(f"\nIndiaPost Tracker  —  {len(article_numbers)} article(s)")
    print("─" * 60)

    auth   = AuthManager(uid, pwd)
    result = fetch_tracking(auth, article_numbers)

    if not result.get("success"):
        print(f"[ERROR] Tracking API: {result.get('message', result)}")
        return

    articles  = result.get("data") or []
    found_ids = {
        a.get("booking_details", {}).get("article_number")
        for a in articles
    }

    print()
    for article in articles:
        booking    = article.get("booking_details", {})
        article_no = booking.get("article_number", "UNKNOWN")
        last_event = get_last_event(article)

        print(f"Article  : {article_no}")
        print(f"Booked   : {booking.get('booked_at', 'N/A')}  "
              f"on  {(booking.get('booked_on') or '')[:10]}")
        print(f"Route    : {booking.get('origin_pincode', '?')} → "
              f"{booking.get('destination_pincode', '?')}")
        print(f"Type     : {booking.get('article_type', 'N/A')}")

        if last_event:
            print(f"Last evt : {format_event(last_event)}")
        else:
            print("Last evt : No events recorded yet")

        del_st = article.get("del_status", {})
        if del_st:
            status_val = del_st.get("del_status", "")
            if status_val:
                print(f"Status   : {status_val.upper()}")

        print()

    # Warn about IDs with no data returned
    for art_no in article_numbers:
        if art_no not in found_ids:
            print(f"[WARN] No data returned for: {art_no}")


# ── CLI entry point ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        ids = sys.argv[1:]
    else:
        raw = input("Enter tracking ID(s) separated by spaces or commas:\n> ").strip()
        ids = [x.strip() for x in raw.replace(",", " ").split() if x.strip()]

    if not ids:
        print("No tracking IDs provided. Exiting.")
        sys.exit(1)

    track_articles(ids)
