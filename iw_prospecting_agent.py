"""
IW Technologies Daily Prospecting Agent
----------------------------------------
Run this script inside Claude Code Remote to process the next batch
of accounts from matt_book_of_business_for_claude.xlsx, cross-reference
with ZoomInfo, check for triggering events, and append results to
the Google Sheet tracker.

Schedule: Run daily at 6am CDT via Claude Code Remote trigger.
"""

import os
import json
import datetime
import anthropic

# ── Configuration ─────────────────────────────────────────────────────────────

ANTHROPIC_MODEL = "claude-sonnet-4-6"
BATCH_SIZE = 14  # accounts per daily run (VP guidance: 12-16)
GOOGLE_SHEET_NAME = "IW Prospecting Tracker - Daily Batches"

# Accounts already processed (update this list as batches complete)
PROCESSED_ACCOUNTS = [
    "skechers", "five below", "coach", "tapestry", "sally beauty",
    "altar'd state", "altardstate", "children's place", "total wine",
    "journeys", "genesco", "zumiez", "spencer", "lids", "lacoste",
    "aldo", "buckle", "levi", "johnston", "container store", "ikea",
    "petco", "kendra scott", "l.l.bean", "anthropologie", "free people",
    "aritzia", "h&m", "shoe carnival", "tory burch", "lilly pulitzer",
    "tiffany", "louis vuitton", "l'occitane", "harbor freight",
    "montblanc", "amer sports", "motiva", "h&r block", "jazzercise",
    "jackson hewitt", "tax service", "sprint nextel", "rent-a-center",
    "cca global", "carpet one", "exxonmobil", "anytime fitness",
    "bridgestone", "imperial oil", "kindercare", "sport clips",
    "verizon wireless", "curves international", "chs inc", "cenex",
    "murphy usa", "carquest", "dfc global", "edible arrangements",
    "cato corporation",
]

# Permanent exclusions
EXCLUDED_ACCOUNTS = [
    "tilly's", "dry goods", "eddie bauer", "charlotte russe", "crocs",
    "rue21", "carters", "lucky", "akira", "sephora", "gamestop",
    "off broadway shoe warehouse", "windsor", "bath and body works",
    "perfumania", "jd sports", "sprint nextel",
]

# Existing customers - do not cold pitch
EXISTING_CUSTOMERS = ["town pump", "belle tire", "buckle"]

# ── System prompt for the prospecting agent ───────────────────────────────────

SYSTEM_PROMPT = """You are a B2B retail IT prospecting agent for IW Technologies, a 50+ year retail IT
services company. IW's 50th anniversary is 2026. Your job is to research retail and adjacent accounts,
find the best IT contacts via ZoomInfo, identify triggering events, and produce outreach-ready prospect
data in JSON format.

IW SERVICES: POS hardware sales (new and certified refurbished), depot services, break/fix support,
new store kitting and staging, low voltage cabling, network equipment.

VERTICAL FIT RULES:
- CONFIRMED FIT: Specialty retail, c-stores (POS only, not pump equipment), auto service,
  grocery/food retail, tax prep offices (H&R Block etc.), rent-to-own, telecom retail stores,
  franchise chains (use franchise-specific pitch)
- FRANCHISE PITCH: Position IW as the vendor corporate recommends to franchisees for warranty
  management, low-cost refurb hardware, and low voltage cabling during corporate-mandated upgrades
- GENERAL RULE: If the account has IT hardware (terminals, pin pads, scanners, printers) then pitch
  a general buy/sell/service POS and network hardware approach
- SKIP ONLY: Pure petroleum refining with no retail stores, pure digital/financial services with no
  physical locations, dead entities

CONTACT STANDARD (2-3 contacts per account):
1. Hardware refresh/depot: VP/Dir IT Infrastructure or Store Systems
2. NSO/kitting: Dir of Store Development or Facilities IT
3. Break/fix/services: IT Operations Manager or Field Tech lead

TRIGGERING EVENTS TO CHECK (via web search and ZoomInfo scoops):
- New CIO/CTO/VP IT appointments (leadership change = vendor review window)
- Acquisitions or rebranding (hardware standardization needed)
- New store format launches or store opening announcements
- POS platform migrations (post-migration hardware support gap)
- Hardware EOL announcements (Ingenico 480, MX925, Engage One original, etc.)
- Earnings calls mentioning technology investment
- RetailDive.com articles about the account

EMAIL STYLE RULES:
- Poke-the-bear style (Josh Braun framework): open with a gap or problem, not a compliment
- Sandler framework: surface the problem, let them self-discover the need
- Max 4 sentences for cold emails, longer for follow-ups when warranted
- No em dashes or en dashes - use plain punctuation only
- Subject lines 3-5 words
- No "hope this finds you well"
- Close with a single diagnostic question, not a meeting ask
- IW social proof: Trader Joe's (10 to hundreds of stores), Ollie's (80 to 650+), Cavender's (50 to 100+)

OUTPUT FORMAT: Return a JSON array of account objects. Each object must have:
{
  "company": "Company Name",
  "stores": "X locations",
  "vertical_fit": "fit|skip|franchise",
  "skip_reason": "Only if skip - brief explanation",
  "priority": "High|Medium|Low",
  "triggered": true|false,
  "trigger_details": "Description of triggering event and source",
  "contacts": [
    {
      "name": "Full Name or TBD",
      "title": "Exact title",
      "role": "Hardware Refresh|Break/Fix|NSO/Kitting",
      "email": "email@company.com or 'Confirm via ZoomInfo' or 'Not yet identified'",
      "phone": "number or status",
      "dnc": false,
      "zi_accuracy": 0,
      "zi_validated": "date or unknown",
      "zi_status": "Confirmed|Confirm via ZoomInfo|Not found"
    }
  ],
  "hardware": "Observed hardware from in-store research or inferred",
  "pitch_angle": "Break/Fix & Depot|Hardware Refresh|NSO/Kitting|Payment Terminal Refresh|General POS & Network",
  "pitch_detail": "1-2 sentence specific angle tied to hardware and signal",
  "email_subject": "3-5 word subject line",
  "email_draft": "Full cold email text ready to send",
  "voicemail": "Voicemail script",
  "talking_points": "3 numbered talking points for live call",
  "signal_source": "RetailDive article URL, ZoomInfo scoop, earnings report, etc.",
  "salesforce_note": "CRM import readiness note"
}
"""

# ── Main agent logic ──────────────────────────────────────────────────────────

def load_account_queue():
    """
    Load and sort the master book of business by store count.
    Returns list of (company_name, stores, my_notes, evan_notes) tuples
    for accounts not yet processed and not excluded.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook("matt_book_of_business_for_claude.xlsx", data_only=True)
        ws = wb.active

        accounts = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            company = row[0]
            my_notes = row[1] or ""
            evan_notes = row[2] or ""
            stores = row[3] or 0
            trade_names = row[4] or ""

            if not company:
                continue

            company_lower = company.lower()

            # Skip excluded and existing customers
            if any(ex in company_lower for ex in EXCLUDED_ACCOUNTS + EXISTING_CUSTOMERS):
                continue

            # Skip already processed
            if any(proc in company_lower for proc in PROCESSED_ACCOUNTS):
                continue

            accounts.append({
                "company": company,
                "stores": stores,
                "my_notes": my_notes,
                "evan_notes": evan_notes,
                "trade_names": trade_names,
            })

        # Sort by store count descending
        accounts.sort(key=lambda x: x["stores"] or 0, reverse=True)
        return accounts

    except Exception as e:
        print(f"Error loading account queue: {e}")
        return []


def run_prospecting_batch(accounts_batch, in_store_hardware: dict):
    """
    Use the Claude API with ZoomInfo MCP and web search to research
    a batch of accounts and return structured prospect data.
    """
    client = anthropic.Anthropic()

    # Build the user prompt with account list and hardware data
    account_list = []
    for acct in accounts_batch:
        hw = in_store_hardware.get(acct["company"].lower(), "Not documented - infer from vertical")
        notes = ""
        if acct["my_notes"]:
            notes += f" Matt's note: {acct['my_notes']}."
        if acct["evan_notes"]:
            notes += f" Evan's note: {acct['evan_notes']}."

        account_list.append(
            f"- {acct['company']} ({acct['stores']} locations){notes} "
            f"In-store hardware: {hw}"
        )

    user_prompt = f"""Today is {datetime.date.today().strftime('%B %d, %Y')}.

Research the following {len(accounts_batch)} accounts from IW Technologies' master book of business.
For each account:
1. Use ZoomInfo to find 2-3 contacts (hardware refresh, NSO/kitting, break/fix roles)
2. Enrich confirmed contacts for email and phone
3. Search the web and RetailDive for triggering events in the last 90 days
4. Apply IW's vertical fit rules - flag skips clearly with reasons
5. Write a cold email draft, voicemail script, and 3 talking points for each non-skip account

ACCOUNTS TO RESEARCH:
{chr(10).join(account_list)}

Return ONLY a valid JSON array. No preamble, no explanation, no markdown code fences.
Each element must follow the output schema exactly."""

    print(f"Running prospecting batch for {len(accounts_batch)} accounts...")
    print(f"Accounts: {[a['company'] for a in accounts_batch]}")

    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
        mcp_servers=[
            {
                "type": "url",
                "url": "https://mcp.zoominfo.com/mcp",
                "name": "zoominfo-mcp"
            }
        ],
        tools=[
            {
                "type": "web_search_20250305",
                "name": "web_search"
            }
        ]
    )

    # Extract the JSON from the response
    result_text = ""
    for block in response.content:
        if block.type == "text":
            result_text += block.text

    # Parse JSON
    result_text = result_text.strip()
    if result_text.startswith("```"):
        result_text = result_text.split("```")[1]
        if result_text.startswith("json"):
            result_text = result_text[4:]
    result_text = result_text.strip().rstrip("```").strip()

    return json.loads(result_text)


def append_to_google_sheet(batch_results: list, sheet_name: str):
    """
    Append batch results to the Google Sheet tracker via the Drive MCP.
    Creates a new dated tab for today's batch.
    """
    today = datetime.date.today().strftime("%Y-%m-%d")
    tab_name = f"{today} Batch"

    # Build CSV rows for the new sheet tab
    headers = [
        "Company", "Stores", "Vertical Fit", "Priority", "Triggered",
        "Trigger Details",
        "Contact 1 Name", "Contact 1 Title", "Contact 1 Role",
        "Contact 1 Email", "Contact 1 Phone",
        "Contact 2 Name", "Contact 2 Title", "Contact 2 Role",
        "Contact 2 Email", "Contact 2 Phone",
        "Contact 3 Name", "Contact 3 Title", "Contact 3 Role",
        "Contact 3 Email", "Contact 3 Phone",
        "Observed Hardware", "Pitch Angle", "Pitch Detail",
        "Email Subject", "Email Draft", "Voicemail Script", "Talking Points",
        "Signal Source", "Salesforce Note",
        "Seq 1: Email", "Seq 2: Call+VM", "Seq 3: Follow-up",
        "Seq 4: LinkedIn", "Seq 5: 30-day", "Seq 6: Triggered",
        "Status / Reply Notes"
    ]

    import csv
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(headers)

    for acct in batch_results:
        contacts = acct.get("contacts", [])
        c = [{} for _ in range(3)]
        for i, contact in enumerate(contacts[:3]):
            c[i] = contact

        def cfield(idx, field):
            return c[idx].get(field, "") if idx < len(contacts) else ""

        row = [
            acct.get("company", ""),
            acct.get("stores", ""),
            acct.get("vertical_fit", ""),
            acct.get("priority", ""),
            "YES" if acct.get("triggered") else "No",
            acct.get("trigger_details", ""),
            cfield(0, "name"), cfield(0, "title"), cfield(0, "role"),
            cfield(0, "email"), cfield(0, "phone"),
            cfield(1, "name"), cfield(1, "title"), cfield(1, "role"),
            cfield(1, "email"), cfield(1, "phone"),
            cfield(2, "name"), cfield(2, "title"), cfield(2, "role"),
            cfield(2, "email"), cfield(2, "phone"),
            acct.get("hardware", ""),
            acct.get("pitch_angle", ""),
            acct.get("pitch_detail", ""),
            acct.get("email_subject", ""),
            acct.get("email_draft", ""),
            acct.get("voicemail", ""),
            acct.get("talking_points", ""),
            acct.get("signal_source", ""),
            acct.get("salesforce_note", ""),
            "Email pending", "Call + VM pending", "Follow-up pending",
            "LinkedIn pending", "30-day pending", "Triggered pending",
            ""
        ]
        writer.writerow(row)

    csv_content = output.getvalue()

    # Upload to Google Drive as a CSV (will convert to Google Sheet)
    client = anthropic.Anthropic()
    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": f"Create a new Google Sheet file titled '{sheet_name} - {tab_name}' with this CSV content:\n\n{csv_content}"
        }],
        mcp_servers=[
            {
                "type": "url",
                "url": "https://drivemcp.googleapis.com/mcp/v1",
                "name": "google-drive-mcp"
            }
        ]
    )

    print(f"Sheet upload response: {response.content[0].text if response.content else 'No response'}")
    return tab_name


def send_email_summary(batch_results: list, tab_name: str):
    """
    Send a summary email via Gmail MCP with today's batch highlights.
    """
    today = datetime.date.today().strftime("%B %d, %Y")
    fit_accounts = [a for a in batch_results if a.get("vertical_fit") != "skip"]
    triggered = [a for a in fit_accounts if a.get("triggered")]
    high_priority = [a for a in fit_accounts if a.get("priority") == "High"]

    summary_lines = [
        f"IW Prospecting - Daily Batch Summary ({today})",
        f"",
        f"Accounts processed: {len(batch_results)}",
        f"IW fit accounts: {len(fit_accounts)}",
        f"High priority: {len(high_priority)}",
        f"Triggered events: {len(triggered)}",
        f"",
        f"TOP ACCOUNTS TO ACT ON TODAY:",
    ]

    # List triggered + high priority first
    priority_accounts = triggered + [a for a in high_priority if a not in triggered]
    for acct in priority_accounts[:5]:
        contacts = acct.get("contacts", [])
        top_contact = contacts[0] if contacts else {}
        summary_lines.append(
            f"  {acct['company']} - {top_contact.get('name', 'TBD')} "
            f"({top_contact.get('email', 'Email TBD')}) - "
            f"{'TRIGGERED: ' + acct.get('trigger_details', '')[:60] if acct.get('triggered') else acct.get('pitch_angle', '')}"
        )

    summary_lines += [
        f"",
        f"Full batch details in Google Sheet: {tab_name}",
        f"",
        f"-- IW Prospecting Agent"
    ]

    email_body = "\n".join(summary_lines)

    client = anthropic.Anthropic()
    client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": f"Send an email to matt@weareiw.com with subject 'IW Prospecting - {today} Batch Ready' and this body:\n\n{email_body}"
        }],
        mcp_servers=[
            {
                "type": "url",
                "url": "https://gmailmcp.googleapis.com/mcp/v1",
                "name": "gmail-mcp"
            }
        ]
    )
    print("Summary email sent.")


def main():
    print(f"IW Prospecting Agent - {datetime.date.today()}")
    print("=" * 60)

    # 1. Load account queue
    queue = load_account_queue()
    if not queue:
        print("No accounts remaining in queue. Exiting.")
        return

    print(f"Accounts remaining in queue: {len(queue)}")

    # 2. Take next batch
    batch = queue[:BATCH_SIZE]
    print(f"Processing batch of {len(batch)} accounts:")
    for a in batch:
        print(f"  - {a['company']} ({a['stores']} stores)")

    # 3. In-store hardware lookup (cross-reference with hardware research sheet)
    # This maps company names to observed hardware from the in-store research sheet
    # Update this dict as new hardware is documented in-store
    in_store_hardware = {
        "sport clips": "POS terminal, pin pad, receipt printer (franchise salon)",
        "jazzercise": "POS tablet, receipt printer (fitness franchise)",
        "jackson hewitt": "Client-facing workstations, receipt printers, tax office terminals",
        "anytime fitness": "POS kiosk, membership terminals (fitness franchise)",
        "bridgestone retail operations": "POS terminals, service management system, receipt printers (auto service)",
        "kindercare education": "Check-in kiosks, administrative workstations",
        "curves international": "POS terminal, membership management kiosk (fitness franchise)",
        "chs inc": "C-store POS, fuel integration terminals, scanners, receipt printers",
        "edible arrangements": "POS terminal, receipt printer (food gift franchise)",
        "verizon wireless": "POS terminals, mobile device demo units, receipt printers (telecom retail)",
        "the cato corporation": "POS terminals, scanners, receipt printers (specialty apparel)",
        "murphy usa": "C-store POS, fuel integration, scanners, receipt printers",
        "h&r block": "Client-facing terminals, workstations, receipt printers (tax prep)",
        "rent-a-center": "POS terminals, customer-facing kiosks, workstations (rent-to-own)",
    }

    # 4. Run the agent
    try:
        results = run_prospecting_batch(batch, in_store_hardware)
        print(f"\nResearch complete. {len(results)} accounts processed.")
    except Exception as e:
        print(f"Error running prospecting batch: {e}")
        raise

    # 5. Append to Google Sheet
    try:
        tab_name = append_to_google_sheet(results, GOOGLE_SHEET_NAME)
        print(f"Results saved to Google Sheet tab: {tab_name}")
    except Exception as e:
        print(f"Error saving to Google Sheet: {e}")

    # 6. Send email summary
    try:
        send_email_summary(results, tab_name)
    except Exception as e:
        print(f"Error sending summary email: {e}")

    print("\nDone.")


if __name__ == "__main__":
    main()
