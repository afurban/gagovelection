import json
import os
import re
import tempfile
from datetime import datetime, timezone
from urllib.request import Request, urlopen
from openpyxl import load_workbook

REPORT_URL = "https://results.sos.ga.gov/cdn/results/09378a07-e6cf-4f66-be7c-ca4aa534f99a/Total%20Votes%20Results_7b88ab32-a923-459d-98e2-33c5324b6428.xlsx"
OUTPUT = "data/2026.json"


def clean_county(value):
    name = str(value or "").strip()
    name = re.sub(r"\s+County$", "", name, flags=re.I)
    # Match the dashboard's conventional spelling.
    if name.lower() == "dekalb":
        return "DeKalb"
    return name


def main():
    request = Request(REPORT_URL, headers={"User-Agent": "Mozilla/5.0 GeorgiaGovernorDashboard/1.0"})
    with urlopen(request, timeout=60) as response:
        content = response.read()
    if len(content) < 1000 or not content.startswith(b"PK"):
        raise RuntimeError("The election report download did not look like an Excel workbook.")

    with tempfile.NamedTemporaryFile(suffix=".xlsx") as temp:
        temp.write(content)
        temp.flush()
        workbook = load_workbook(temp.name, data_only=True, read_only=True)
        if "County Results" not in workbook.sheetnames:
            raise RuntimeError("The downloaded workbook does not contain the expected 'County Results' worksheet.")
        sheet = workbook["County Results"]
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if not header or [str(x or "").strip() for x in header[:7]] != ["County", "Office Name", "Contest ID", "Ballot Name", "Choice ID", "Party", "Total"]:
            raise RuntimeError("The County Results worksheet has an unexpected column layout; stopping to avoid bad data.")

        results = {}
        for row in rows:
            if len(row) < 7:
                continue
            county, office, _, ballot_name, _, party, total = row[:7]
            if str(office or "").strip().lower() != "governor":
                continue
            county_name = clean_county(county)
            party_name = str(party or "").strip().lower()
            ballot = str(ballot_name or "").strip().lower()
            if not county_name:
                continue
            item = results.setdefault(county_name, {"r": 0, "d": 0, "o": 0})
            votes = int(total or 0)
            if party_name in ("rep", "republican"):
                item["r"] += votes
            elif party_name in ("dem", "democratic"):
                item["d"] += votes
            elif ballot not in ("total votes", "ballots cast"):
                item["o"] += votes

    # Only publish counties with actual candidate votes. A zero-filled early report
    # must not make all 159 counties appear as reported.
    results = {county: votes for county, votes in results.items() if votes["r"] + votes["d"] + votes["o"] > 0}
    if not results:
        print("The report currently contains zero Governor candidate votes. Keeping the dashboard's existing 2026 data unchanged.")
        return

    payload = {
        "source": REPORT_URL,
        "updated_utc": datetime.now(timezone.utc).isoformat(),
        "results": results,
    }
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as output:
        json.dump(payload, output, ensure_ascii=False, indent=2)
        output.write("\n")
    print(f"Saved 2026 Governor results for {len(results)} counties to {OUTPUT}.")


if __name__ == "__main__":
    main()
