"""
indiapost_api.py  —  Safe, read-only wrapper around the IndiaPost Production API.

This module ONLY reads data. It does not modify indiapost_tracker.py or
indiapost_report.py and is safe to import alongside them.

Exposes:
  get_tracking_data(ids)   -> dict[article_number, raw_api_dict]
  get_receipt_json(ids)    -> dict[article_number, receipt_dict]
  get_event_history(id)    -> list[event_dict]

All network calls go through indiapost_tracker.AuthManager / fetch_tracking
so credentials and token refresh are handled in one place.
"""

import math
import time as _time
from indiapost_tracker import AuthManager, fetch_tracking, CLIENT_ID, CLIENT_SECRET

BATCH_SIZE = 20   # keep small — API silently drops tracking_details on large batches


def _auth() -> AuthManager:
    a = AuthManager(CLIENT_ID, CLIENT_SECRET)
    a.get_token()
    return a


def get_tracking_data(ids: list[str],
                      auth: AuthManager = None,
                      verbose: bool = True) -> dict[str, dict]:
    """
    Fetch full tracking data for a list of article numbers.
    Returns dict keyed by article_number -> raw API article object.
    Handles batching, timeouts (3 retries), and null-tracking_details retry.
    """
    if not auth:
        auth = _auth()

    result_map: dict[str, dict] = {}
    batches = math.ceil(len(ids) / BATCH_SIZE)

    for i in range(batches):
        batch = ids[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        if verbose:
            print(f"  Batch {i+1}/{batches}  ({len(batch)} articles)...", end=" ", flush=True)

        resp = {}
        for attempt in range(3):
            try:
                resp = fetch_tracking(auth, batch)
                break
            except Exception as e:
                if attempt < 2:
                    if verbose:
                        print(f"[retry {attempt+2}]...", end=" ", flush=True)
                    _time.sleep(5)
                else:
                    if verbose:
                        print(f"[FAILED: {e}]")

        for art in (resp.get("data") or []):
            num = (art.get("booking_details") or {}).get("article_number", "")
            if num:
                result_map[num] = art

        if verbose:
            print("done")

    # Retry any that returned with null tracking_details
    nulls = [id_ for id_ in ids
             if id_ in result_map and not result_map[id_].get("tracking_details")]
    if nulls:
        if verbose:
            print(f"  Retrying {len(nulls)} with null tracking_details...")
        for id_ in nulls:
            try:
                resp = fetch_tracking(auth, [id_])
                for art in (resp.get("data") or []):
                    num = (art.get("booking_details") or {}).get("article_number", "")
                    if num and art.get("tracking_details"):
                        result_map[num] = art
            except Exception:
                pass

    return result_map


def get_receipt_json(ids: list[str],
                     auth: AuthManager = None) -> dict[str, dict]:
    """
    Returns a clean receipt dict per article, suitable for JSON export or
    later printing. Pulls all available fields from booking_details +
    full event history.

    Receipt dict keys:
      article_number, article_type, booked_at, booked_on,
      origin_pincode, destination_pincode, delivery_location,
      delivery_confirmed_on, tariff, del_status,
      events: [ {date, time, office, event}, ... ]
    """
    raw = get_tracking_data(ids, auth=auth)
    receipts = {}
    for num, art in raw.items():
        b = art.get("booking_details") or {}
        events = art.get("tracking_details") or []
        receipts[num] = {
            "article_number":       b.get("article_number", num),
            "article_type":         b.get("article_type", ""),
            "booked_at":            b.get("booked_at", ""),
            "booked_on":            (b.get("booked_on") or "")[:10],
            "origin_pincode":       b.get("origin_pincode", ""),
            "destination_pincode":  b.get("destination_pincode", ""),
            "delivery_location":    b.get("delivery_location", ""),
            "delivery_confirmed_on":(b.get("delivery_confirmed_on") or "")[:10],
            "tariff":               b.get("tariff", ""),
            "del_status":           (art.get("del_status") or {}).get("del_status", ""),
            "events": [
                {
                    "date":   (e.get("date") or "")[:10],
                    "time":   e.get("time", ""),
                    "office": e.get("office", ""),
                    "event":  e.get("event", ""),
                }
                for e in events
            ],
        }
    return receipts


def get_event_history(article_id: str,
                      auth: AuthManager = None) -> list[dict]:
    """Return the full event list for a single article (newest first)."""
    data = get_tracking_data([article_id], auth=auth, verbose=False)
    art  = data.get(article_id, {})
    return art.get("tracking_details") or []
