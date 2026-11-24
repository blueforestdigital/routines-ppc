#!/usr/bin/env python3
"""Weekly Negative Keyword Analysis — 2026-11-24 (data: Nov 17–23, 2026).

Usage: process_nkw_2026-11-24.py <data_dir> <out_xlsx> [summary_json]
<data_dir> holds one <TabName>.json per account: the raw Windsor.ai google_ads rows
(search_term, campaign, ad_group, impressions, clicks, spend, conversions, conversion_value).
"""
import json, os, re, sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

DATA_DIR, OUT = sys.argv[1], sys.argv[2]
SUMMARY_JSON = sys.argv[3] if len(sys.argv) > 3 else None
DATE_RANGE = "Nov 17–23, 2026"

TAB_ORDER = ["ASRV", "Griffin", "Velo", "Sunrise", "TriActive", "LEARN", "Eurolog", "Spotlight",
             "CQC", "EcoPeaCo CA", "EcoPeaCo US", "GoGather", "RCM", "Ripco", "We Buy"]
DISPLAY = {"ASRV": "ASRV", "Griffin": "Griffin Funding", "Velo": "Velo Property Solutions",
           "Sunrise": "Sunrise Yacht Products", "TriActive": "TriActiveUSA", "LEARN": "LEARN Behavioral",
           "Eurolog": "Eurolog Packing Group", "Spotlight": "Spotlight.Vegas", "CQC": "Custom Quality Covers",
           "EcoPeaCo CA": "EcoPeaCo Canada", "EcoPeaCo US": "EcoPeaCo USA", "GoGather": "GoGather",
           "RCM": "Renewable Credit Management", "Ripco": "Ripco Products", "We Buy": "We Buy Exotics"}
# Accounts with nothing to analyze, and why.
NO_DATA = {
    "Velo": "No data returned — skipping (account 135-061-8133 is not connected in Windsor.ai)",
    "CQC": "No data returned — skipping (no activity in the last 7 days; account is marked Paused in Client Budgets)",
    "RCM": "No data returned — skipping (Windsor.ai returned zero rows for the last 7 days, even unfiltered)",
}
# CPA used for spend-vs-CPA judgments. From the client's CLAUDE.md where it gives one;
# otherwise max($50 default, observed 7-day cost/conversion) — see cpa_for().
TARGET_CPA = {"Eurolog": 500, "Sunrise": 150, "TriActive": 150, "LEARN": 200}

E, C, K = "ELIMINATE", "CONSIDER", "COMPETITOR"
COMP_NOTE = " — review only; may be intentional conquesting"


def load(key):
    path = os.path.join(DATA_DIR, key + ".json")
    if not os.path.exists(path):
        return []
    agg = {}
    for r in json.load(open(path)):
        t = r.get("search_term")
        if not t:  # PMax / Display / Demand Gen rows carry no search term
            continue
        if not (r["clicks"] > 0 or r["spend"] >= 1.0):  # Step 3a threshold
            continue
        x = agg.setdefault(t, {"search_term": t, "campaigns": [], "impressions": 0, "clicks": 0,
                               "spend": 0.0, "conversions": 0.0, "conversion_value": 0.0})
        if r["campaign"] not in x["campaigns"]:
            x["campaigns"].append(r["campaign"])
        for f in ("impressions", "clicks", "spend", "conversions", "conversion_value"):
            x[f] += r[f] or 0
    return list(agg.values())


def cpa_for(key, terms):
    if key in TARGET_CPA:
        return TARGET_CPA[key], "CLAUDE.md"
    spend = sum(t["spend"] for t in terms)
    conv = sum(t["conversions"] for t in terms)
    observed = spend / conv if conv else 0
    return (max(50.0, observed), "observed 7-day cost/conv" if observed > 50 else "$50 default")


# ── per-client config: brand regex, competitor regex, ordered (regex, category, reason) rules ──
JOBS = (r"\b(jobs|careers?|salary|hiring|internship)\b", E, "Job-seeking query — no buyer intent")

GRIFFIN_COMP = (
    r"athena mortgage|trustmark|novus bank|salt funding|churchill|amerisave|\bbmo\b|5 3 bank|wells fargo|"
    r"navy (federal|credit)|nation capital|ally bank|ladera|veterans first|triad|fairbridge|homesimply|anchor home|"
    r"vanderbilt|\baven\b|servion|streit|global credit union|21st ?mortgage|21 (first )?mortgage|mortgage 21|td bank|"
    r"superior mortgage|tr[ao]nchi|union home mortgage|agave|capital fund 1|caprate|quontic|bell bank|nexa|"
    r"richmond mortgage|backflip|jmj financial|integrity mortgage|morgan stanley|clear mortgage capital|annie mac|"
    r"cardinal financial|sistar|third federal|newsilver|highland mortgage|opportunity mortgage|tomo mortgage|"
    r"pb financial|truewealth|dream america|goldwater|bluedog|brightstar|clarity capital|\baxos\b|supreme lending|"
    r"all american funding|mld mortgage|caplink|oak mortgage|homeside|e mortgage capital|emortgage|better mortgage|"
    r"rockland trust|truist|fairway|capital home loans|\bchase\b|arbor trust|evolve funding|stormfield|splitero|"
    r"investor iq|hamilton mortgage|nfc lending|\bcmg\b|blue sky financial|silverton|merrill|century 21|flat iron|"
    r"rocket|kennedy funding|executive home loans|fynn financial|stock ?yards?|west capital|lendgo|neo home loans|"
    r"green ?leaf lending|american lending llc|a & d mortgage|lendex|rivers edge|southeast funding|ownify|"
    r"harrison financial|more life mortgage|freedom ?(finance|mortgage)|uif corporation|usaa|smart homebuyers|"
    r"cornerstone mortgage|fairfield mortgage|walley|silverstone|alecko|alliant|lmcu|planet (mortgage|home)|"
    r"wings financial|longhorn loans|champions funding|dedham|gentry funding|modern mortgage|c2 financial|"
    r"loan care|coast to coast funding|calcon|cedar creek|lake superior lending|first heritage|ltv lender|"
    r"ark la tex|amerihome|statewide mortgage|desert ridge|liam and anderson|webster bank|onity|crown home|"
    r"american heritage|1st franklin|medallion|sentinel mortgage|prime finance|gfa mortgage|creative owners club|"
    r"guaranteed mortgage")
GRIFFIN_GENERIC_NAV = (
    r"^(credit unions?( near me)?|banks?( near me)?|delaware banks|credit union loans|"
    r"credit unions in dallas texas|cdfi credit unions near me|private equity|accredited investors|infinite banking|"
    r"secondary mortgage market|lending and borrowing|real estate closing|mortgage credit report|capital|"
    r"biggerpockets.*|cash flowing deals|foreclosure|multifamily)$")

ASRV_COMP = (
    r"darc sport|vuori|gymshark|raised by wolves|young ?la|\bjox\b|under armou?r|ten ?thousand|\b10000\b|salomon|"
    r"mayweather|lululemon|all ?saints|tracksmith|municipal|\btlf\b|castore|rose endurance|\bhaus\b|"
    r"tailored athlete|no ?bull|ryderwear|real ?tree|bandit|rhone|path projects|nike|dead ?last|alphalete|\buvu\b|"
    r"three nails|trained by jp|sia wear|wolaco|fabletics|new balance|hwpo|runnractive|\balo\b|unknown runner|"
    r"project rock|nero clothing|ons clothing|team 247|represent|thule|raw nutrition|pme legend|\blegends\b|kuiu|"
    r"off white|satisfy|killcrew|virus|vanquish|safetti|bum apparel|running ?warehouse|adidas|govx|021 originals|"
    r"hoka|koru athletics|aybl|rogue fitness|h&m|\bmous\b|loiter|cuts clothing|stubble and co|o neill|puma|"
    r"boohooman|arcteryx|barbell brigade|squat wolf|malbon|\bon running|prada|patencials|9forty|"
    r"every day better|hyrox")

LEARN_COMP = (
    r"circle city|hope ?bridge|little star|perfect child|butterfly effects|boston behavior learning|anna aba|"
    r"cortica|stepping stones|lighthouse|above and beyond|brighter strides|era aba|elevate spectrum|"
    r"constellation kids|ally behavior|helping hands|bright day|alphabet soup|shapiro aba|applied abc|"
    r"kaleidoscope|beacon services|grateful care|proud moments|verbal beginnings|golden steps|"
    r"autism ?faith ?network|\bcabi\b|pbs aba|positive behavior supports|renoxx|logan center|achieve aba|"
    r"blue gems|centria|ivy aba|magnet aba|heartlinks|bright bridge|kids first services|aba centers? of america|"
    r"arms wide open|autism learning partners|total care aba|tiny steps|stride autism|\bicbd\b|"
    r"autism care partners|yes aba|autism learning world")

CFG = {
    "ASRV": {
        "brand": r"asrv|a s r v|arsv|asvr|astrv|asrw|asrbv|asrb|asrp|aesthetic revolution",
        "brand_hybrid": (r"discount code|coupon|promo|reddit|review", "Brand + discount/review modifier — likely deal-seeker or researcher; review match type"),
        "comp": ASRV_COMP,
        "rules": [
            (r"macbook|潮|nate diaz|kobe|fishnet|tj maxx|^youth |why are running vests", E, "Unrelated product or non-buyer query — no fit with ASRV technical activewear"),
            (r"^(black friday deals|best early black friday sales)$", E, "Generic deal-hunting query with no apparel intent"),
            (r"^(srv|rsrv|arv|arvs|asvp) clothing$|^rsva$|^asr shorts$", C, "Possible ASRV misspelling or a different brand — verify before acting"),
            (r"swim trunks|boardshorts|snowboard|triathlon|jiu jitsu|\bbjj\b|basketball", C, "Adjacent sport/product category outside core training + streetwear range"),
            (r"crossbody|cross body|fanny pack|sling bag|macbook|luxury backpacks", C, "Accessory query — confirm ASRV carries a matching product"),
        ],
    },
    "Griffin": {
        "brand": r"griffin|griffith funding",
        "brand_hybrid": (r"review|login|phone|complaint|legit", "Brand + navigational/research modifier — likely existing customer or researcher"),
        "comp": GRIFFIN_COMP,
        "rules": [
            (r"\bnaca\b|naga program", E, "NACA nonprofit homebuyer program — not a Griffin product; low-income program audience"),
            (r"cdfi|section 8|calhfa|ihda|usda|housing finance|finance authority|home ?buyer programs?|home buying programs|"
             r"law enforcement|cancer patients|government (backed|repossessed)|the home program|house programs|"
             r"homegrown project|red loan program|renewed hope", E, "Government/assistance-program seeker — outside DSCR / Non-QM / bank-statement focus"),
            (r"owner financ|seller financ|owner carry|rent to own|lease to own|contract for deed|(?<!refinance )land contract|"
             r"buy a house without a bank|take over someones mortgage|payment plan", E, "Seeking owner/seller financing or rent-to-own — not a lender product"),
            (r"for sale|zillow|properties on market|^foreclosed homes|selling foreclosed|bank owned|apartment buildings", E, "Property-listing search — shopping for a home, not a loan"),
            (r"hotel financing", E, "Commercial loan intent — Griffin does not offer commercial loans"),
            (r"business loan", E, "Business-loan intent — not a mortgage product Griffin offers"),
            (r"cash advance|quick loan|credit ninja|loan company fayetteville", E, "Small personal / cash-advance loan intent — far below the $100K minimum loan"),
            (r"tanya sushkova|geza rakoczi|abel tejeda|camilo escandon|60dbo|nmls id", E, "Navigational lookup for an individual or licence ID — no loan intent"),
            (GRIFFIN_GENERIC_NAV, E, "Generic banking/finance navigation or research — no mortgage-application intent"),
            (r"what is a nmls|what does mortgage mean|fed expected|when will interest rate drop|how to become a real estate investor|"
             r"how to underwrite|how to start real estate business|\breo\b|flex cash|short sale in real estate|"
             r"key questions to ask|purchase agreement for house|how to find investors|private investors near me|"
             r"local investors|fanniemae com|when is the best time to buy|steps to buying a house|"
             r"what makes mortgage rates|rate news|news daily|rate chart|rate tracker|rate trends|prime rate|pros and cons", E,
             "Informational / market-research query with no loan-application intent"),
            (r"\b(bank|credit union)\b(?! (statement|account))", K, "Named bank/credit union"),
            (r"under 100k", C, "Below Griffin's $100K minimum loan amount"),
            (r"mobile home|manufactured|modular|chattel|clayton", C, "Manufactured/mobile-home financing — confirm Griffin product fit"),
            (r"\bland\b|construction|spec house|build", C, "Land/construction financing — confirm Griffin product fit"),
            (r"calculator|calculation", C, "Calculator/tool seeker — low application intent"),
            (r"\bfha\b|203k|freddie mac|fannie mae", C, "FHA/agency loan query — outside priority DSCR / Non-QM / bank-statement focus"),
            (r"apartment|multifamily|mixed use", C, "Multifamily/mixed-use — confirm eligibility (no commercial loans)"),
            (r"(loans?|get a loan).*(no|without).*(income|proof|check stubs|employment)|loan that goes by your income", C,
             "Ambiguous — could be a no-doc mortgage seeker (relevant) or a personal-loan seeker"),
            (r"foreclosure|christian mortgage|requisitos|en español|healthcare workers|nurses", C, "Niche/distressed or non-core audience — review landing-page fit"),
            (r"^(?!.*(dscr|dsr|bank statement|self employed|1099|asset))((what|how|can|do|does|is|why|when|where|i)\b)", C, "Informational modifier — research-stage intent; review if worth the CPC"),
        ],
    },
    "Sunrise": {
        "brand": r"sunrise yacht|sunriseyacht",
        "comp": None,
        "rules": [(r"^multihullnets$", C, "Looks like the client's own domain (multihullnets.com) serving in a Non-Brand campaign — possible brand leakage; verify")],
    },
    "TriActive": {
        "brand": r"triactive ?usa|^triactive$",
        "comp": None,
        "rules": [
            (r"triactive america", C, "Brand-adjacent — 'TriActive America' may be a separate company; verify this is client brand traffic"),
            (r"elliptical|stationary (exercise )?bike", C, "Single-machine query — may be a consumer shopping for home use rather than a parks/corrections buyer"),
            (r"^best ", C, "'Best' modifier — research-stage and possibly consumer intent"),
        ],
    },
    "LEARN": {
        "brand": r"learn behavioral|little leaves|sparks? aba|total spectrum|autism spectrum therapies",
        "brand_hybrid": (r"review|reddit|complaint", "Brand + review modifier — family researching the provider; review only"),
        "comp": LEARN_COMP,
        "rules": [
            (r"\brbt\b|bacb", E, "Professional-credential query (RBT/BACB) — a practitioner or job-seeker rather than a family seeking care"),
            (r"how to do aba therapy", E, "Practitioner/how-to query — not a family looking for a care provider"),
            (r"fort lauderdale", E, "Location outside LEARN's advertised markets (served by a Virginia campaign)"),
            (r"meetup|social groups|respite care", E, "Family seeking community or respite support — a different service than ABA therapy"),
            (r"grants?|medicaid|waiver", C, "Family looking for funding resources — may still need a provider; review"),
            (r"school|daycare", C, "Family looking for school/daycare placement — adjacent to LEARN's school-based support; review"),
            (r"what is aba|explain how aba|when is aba|without autism diagnosis|how to treat|early start denver|speech therapy|music therapy", C,
             "Early-research or adjacent-therapy query — family may not yet be seeking an ABA provider"),
            (r"\d{3,} .*\b(blvd|rd|st|ave)\b", C, "Street-address lookup — confirm whether this is a LEARN location or another provider's"),
            (r"autizum|^aba ft$|autism companies|autism association|autism society|autism resource", C, "Unclear intent or general autism-resource query — review"),
        ],
    },
    "Eurolog": {
        "brand": r"eurolog|epgna|hybag|hyblanket|hypad|temcore",
        "comp": None,
        "conservative": True,
        "rules": [
            (r"foam board|insulat(ion|ed) panels|insulating a", E, "Construction-insulation query (container building/DIY) — client sells cargo thermal liners, not building insulation"),
            (r"con+ex.*insulation|seacan insulation|^container insulation$", C, "Container insulation — may be container-home/storage intent rather than cargo protection; review"),
            (r"electric heating|heated blanket", C, "Powered heating blanket — client sells passive thermal blankets; review fit"),
        ],
    },
    "Spotlight": {
        "brand": r"spotlight",
        "comp": r"\bvegas ?com\b|lasvegas ?com|tix ?4|ticketmaster|stubhub|viator|groupon|vivid ?seats|seatgeek|tripadvisor|"
                r"expedia|get ?your ?guide|klook|travelzoo|best ?of ?vegas|todaytix|goldstar|vegas4locals|cheaptickets|tiqets",
        "rules": [
            (r"\b(auditions?|casting)\b", E, "Performer/audition query — not a ticket buyer"),
            (r"\b(menu|restaurants?|buffet|parking)\b", E, "Dining/parking query — no show-ticket intent"),
            (r"^(?!male reviews)(.*)\b(reviews?|reddit|wiki|youtube|netflix|trailer|lyrics|soundtrack)\b", C, "Research/media modifier — may not be ready to buy tickets"),
            (r"\b(hours|address|phone|dress code)\b", C, "Venue-logistics query — likely already holds tickets or just planning"),
            (r"\bfree\b", C, "'Free' modifier — low purchase intent"),
            (r"near me", C, "'Near me' query — may be outside Las Vegas; check geo"),
            ("__HOTEL__", C, "Hotel/venue navigational query with no conversions — not clearly show-ticket intent"),
        ],
    },
    "EcoPeaCo CA": {
        "brand": r"eco ?pea",
        "comp": r"coteries?|parasol|hello bello|honest company|kirkland|joone|omaiki|royale|motherease|nora's nursery|"
                r"freestyle|boundless|amore diapers|sam's club|baby gold",
        "rules": [
            (r"reusable|cloth|clothes diapers|diaper flats", E, "Cloth/reusable diaper intent — client sells disposable diapers"),
            (r"how often", E, "Parenting how-to query — no purchase intent"),
            (r"healthy baby", C, "May refer to the Healthybaby brand rather than a generic term — review"),
            (r"couches?\b", C, "French-language query — confirm a French landing experience exists"),
            (r"black friday|deals", C, "Deal-seeker modifier — review margin fit"),
        ],
    },
    "EcoPeaCo US": {
        "brand": r"eco ?pea",
        "comp": r"amazon|target|coterie|parasol|hello bello|honest company|kirkland",
        "rules": [
            (r"reusable|cloth", E, "Cloth/reusable diaper intent — client sells disposable diapers"),
            (r"eco wave|eco green", C, "Possibly another brand's product name — verify"),
        ],
    },
    "GoGather": {
        "brand": r"go ?gather",
        "comp": r"brightspot|bishop mccann|shift alt|gsi tours|cohera|^ce go$|meritage events|metro events|gather event platform|san diego event pros",
        "rules": [
            (r"cheap venues", E, "Budget venue-rental query — consumer/private-event intent, not corporate event management"),
            (r"family feud", E, "Looking for an entertainment/game vendor — not event management services"),
            (r"event (spaces?|venue)|meeting rooms|auditorium|convention center", C, "Venue-rental intent — may want a space, not an event-management agency; review"),
            (r"team building", C, "Team-building activity search — adjacent to corporate events; review"),
        ],
    },
    "Ripco": {
        "brand": r"rip ?co",
        "comp": None,
        "rules": [(r"correctional industries", C, "'Correctional Industries' is the name of state prison-labor furniture programs — may be navigational; review")],
    },
    "We Buy": {
        "brand": r"we ?buy ?exotics",
        "brand_hybrid": (r"review|legit|scam|1040 calle", "Brand + trust/address modifier — researcher or visitor lookup; review rather than exclude"),
        "comp": r"il+us+o|^iusso$|j ?(&|and|n)? ?s ?auto ?haus|exotic car trader|exoticautosales",
        "rules": [
            (r"supra mk4|prices on|aston martin sale|listings", E, "Buy-side/price-research query — searcher wants to buy or browse, not sell a car"),
            (r"1040 calle", C, "Client street-address lookup — likely navigation to the office; review"),
            (r"dealers?$|exotic car (company|broker)|consignment", C, "Dealer/broker/consignment query — could be a buyer rather than a seller; review"),
            (r"^how to", C, "Informational modifier — research-stage seller"),
        ],
    },
}
HOTEL = re.compile(r"^(the )?(mgm grand|excalibur|luxor|planet hollywood|harrahs|caesars palace|mandalay bay|wynn|"
                   r"venetian|bellagio|aria|stratosphere|the strat|linq|flamingo|paris|new york new york|rio|westgate)"
                   r"( hotel| las vegas| vegas| casino| and casino| resort)*$")


def categorize(t, key, cpa, cpa_src):
    cfg = CFG[key]
    term, spend, conv = t["search_term"].lower(), t["spend"], t["conversions"]
    cat = reason = None
    is_brand = bool(re.search(cfg["brand"], term))
    if is_brand:
        bh = cfg.get("brand_hybrid")
        if bh and re.search(bh[0], term):
            cat, reason = C, bh[1]
    else:
        if cfg.get("comp"):
            m = re.search(cfg["comp"], term)
            if m:
                cat, reason = K, f"References another brand ('{m.group(0).strip()}')" + COMP_NOTE
        if not cat:
            for pat, c, why in [JOBS] + cfg["rules"]:
                if pat == "__HOTEL__":
                    hit = conv == 0 and spend >= 5 and HOTEL.match(term)
                else:
                    hit = re.search(pat, term)
                if hit:
                    cat, reason = c, why + (COMP_NOTE if c == K else "")
                    break
    # Guard: a term with conversions can never be ELIMINATE.
    if cat == E and conv > 0:
        cat, reason = C, f"Drove {conv:g} conversion(s) but intent looks misaligned — {reason[0].lower() + reason[1:]}"
    # Spend-vs-CPA review for on-target, zero-conversion terms (never ELIMINATE on spend alone).
    if not cat and conv == 0 and spend >= cpa and not is_brand:
        mult = spend / cpa
        band = "over 3×" if mult > 3 else "1–3×"
        cat = C
        reason = (f"On-target but ${spend:,.0f} spent with zero conversions — {band} the ${cpa:,.0f} CPA benchmark "
                  f"({cpa_src}); review bids/match type, allowing for conversion lag")
    # Rule-based CONSIDER flags under $1.00 with no conversions are dropped as not worth review time
    # (brand hybrids and ELIMINATE patterns are kept at any spend).
    if cat == C and spend < 1.0 and conv == 0 and not is_brand:
        return None, None
    if cat == E and cfg.get("conservative") and spend < 3:
        cat = C
        reason = "Low-volume account, minimal spend — " + reason[0].lower() + reason[1:]
    return cat, reason


# ── Excel ──
FILLS = {E: PatternFill("solid", fgColor="C00000"), C: PatternFill("solid", fgColor="FFFF00"),
         K: PatternFill("solid", fgColor="ED7D31")}
CAT_FONT = {E: Font(name="Arial", bold=True, color="FFFFFF"), C: Font(name="Arial", bold=True, color="000000"),
            K: Font(name="Arial", bold=True, color="FFFFFF")}
BASE, BOLD = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
HEADERS = ["Category", "Search Term", "Matched Campaigns", "Impressions", "Clicks", "Spend",
           "Conversions", "Conv. Value", "Reason"]
WIDTHS = [14, 44, 46, 13, 9, 12, 13, 13, 95]
ORDER = {E: 0, C: 1, K: 2}


def write_sheet(ws, key, flagged, total, cpa_note):
    ws.merge_cells("A1:I1")
    ws["A1"] = f"{DISPLAY[key]} — Negative Keyword Report | Last 7 Days: {DATE_RANGE}"
    ws["A1"].font = Font(name="Arial", size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left")
    last = 8 + max(len(flagged), 1)
    rng = f"$A$9:$A${last}"
    block = [("Total Search Terms Analyzed", total), ("Terms to Eliminate", f'=COUNTIF({rng},"ELIMINATE")'),
             ("Terms to Consider", f'=COUNTIF({rng},"CONSIDER")'), ("Competitor Terms", f'=COUNTIF({rng},"COMPETITOR")')]
    for i, (m, v) in enumerate(block, start=3):
        ws.cell(i, 1, m).font = BOLD
        ws.cell(i, 2, v).font = BASE
        ws.cell(i, 2).alignment = Alignment(horizontal="left")
    ws.cell(3, 3, "Unique terms with ≥1 click or ≥$1.00 spend (Windsor.ai google_ads, last_7d), aggregated across campaigns. CONSIDER flags under $1.00 spend are omitted.").font = Font(name="Arial", size=9, italic=True)
    if cpa_note:
        ws.cell(4, 3, cpa_note).font = Font(name="Arial", size=9, italic=True)
    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(8, c, h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E4057")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    if not flagged:
        ws.cell(9, 2, NO_DATA.get(key, "No terms flagged this week.")).font = Font(name="Arial", size=10, italic=True)
    for r, t in enumerate(flagged, start=9):
        vals = [t["cat"], t["search_term"], ", ".join(t["campaigns"]), t["impressions"], t["clicks"],
                round(t["spend"], 2), round(t["conversions"], 2), round(t["conversion_value"], 2), t["reason"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell.font = BASE
        ws.cell(r, 1).fill, ws.cell(r, 1).font = FILLS[t["cat"]], CAT_FONT[t["cat"]]
        ws.cell(r, 1).alignment = Alignment(horizontal="center")
        ws.cell(r, 4).number_format = "#,##0"
        ws.cell(r, 6).number_format = ws.cell(r, 8).number_format = "$#,##0.00"
        ws.cell(r, 7).number_format = "0.##"
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A9"


def main():
    wb = Workbook()
    wb.remove(wb.active)
    summary = {}
    for key in TAB_ORDER:
        terms = [] if key in NO_DATA else load(key)
        flagged, note = [], None
        if terms:
            cpa, src = cpa_for(key, terms)
            note = f"CPA benchmark used for spend review: ${cpa:,.0f} ({src})"
            for t in terms:
                cat, reason = categorize(t, key, cpa, src)
                if cat:
                    flagged.append({**t, "cat": cat, "reason": reason})
            flagged.sort(key=lambda t: (ORDER[t["cat"]], -t["spend"]))
        write_sheet(wb.create_sheet(key), key, flagged, len(terms), note)
        cnt = {c: sum(1 for t in flagged if t["cat"] == c) for c in (E, C, K)}
        summary[key] = {"client": DISPLAY[key], "total": len(terms), **cnt, "no_data": NO_DATA.get(key),
                        "spend": {c: round(sum(t["spend"] for t in flagged if t["cat"] == c), 2) for c in (E, C, K)},
                        "top_eliminate": [{"term": t["search_term"], "spend": round(t["spend"], 2), "reason": t["reason"]}
                                          for t in flagged if t["cat"] == E][:5]}
        print(f"{key:12} total={len(terms):5} E={cnt[E]:3} C={cnt[C]:3} K={cnt[K]:3}")
    wb.save(OUT)
    if SUMMARY_JSON:
        json.dump(summary, open(SUMMARY_JSON, "w"), indent=1, ensure_ascii=False)


if __name__ == "__main__":
    main()
