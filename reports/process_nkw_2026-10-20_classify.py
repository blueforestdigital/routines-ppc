import json,re,collections
S='/tmp/claude-0/-home-user-routines-ppc/51198b25-6b46-505d-8ecb-1c53364ebeb3/scratchpad/'

def agg_raw(path,codes):
    agg=collections.OrderedDict()
    for l in open(path):
        l=l.strip()
        if not l or l.startswith('#'): continue
        t,c,imp,clk,sp,cv,val=l.split('|')
        clk=float(clk);sp=float(sp)
        if clk==0 and sp<1: continue
        a=agg.setdefault(t,dict(term=t,camps=[],imp=0,clicks=0,spend=0,conv=0,val=0))
        cn=codes.get(c,c)
        if cn not in a['camps']: a['camps'].append(cn)
        a['imp']+=int(imp);a['clicks']+=clk;a['spend']+=sp;a['conv']+=float(cv);a['val']+=float(val)
    return sorted(agg.values(),key=lambda x:-x['spend'])

def load_json(n): return json.load(open(S+n))

def parse_small(section):
    lines=[];on=False
    for l in open(S+'small_raw.txt'):
        if l.startswith('#'):
            on=(l.strip()=='#'+section); continue
        if on: lines.append(l)
    open(S+f'_{section}.txt','w').writelines(lines)
    return S+f'_{section}.txt'

def R(*a): return re.compile('|'.join(a),re.I)

def band(x,bench):
    return ('CONSIDER',f"No conversions on ${x['spend']:,.0f} spend ({x['spend']/bench:.1f}x the ${bench} CPA benchmark); intent looks on-target, so review bids/match type rather than negating")

# ---------------- ASRV ----------------
A_BRAND=R(r"\basrv\b",r"\ba s r v\b",r"aesthetics? revolution",r"relentless pursuit",r"\baserv\b",r"asr shorts",r"\basrvp\b",r"\bsasrv\b")
A_COMP=R(r"gymshark",r"vuori",r"lululemon",r"\blulu\b",r"young ?la\b",r"roark",r"ten ?thousand",r"tenthousand",r"10 ?thousand",r"\balo\b",r"rhone",r"satisfy",r"bandit",r"\bnike",r"adidas",r"adizero",r"under armo",r"\bhaus\b",r"dfyne|defyne",r"\blrx\b",r"\boner\b",r"\brepresent\b|247 represent",r"aybl",r"darc sport",r"mileoff",r"alphalete",r"buffbunny",r"ryzon",r"ryderwear",r"vanquish fitness",r"\b2xu\b",r"brooks",r"\bhoka\b",r"janji",r"tracksmith",r"sweaty betty",r"set active",r"fabletics",r"\baer\b",r"\btumi\b",r"dsptch",r"kilometer studios",r"nordstrom",r"all saints",r"\bquince\b",r"uniqlo",r"\bzara\b",r"new balance",r"\bpuma\b",r"reebok",r"salomon",r"arc'?teryx",r"north face",r"castore",r"\bciele\b",r"saysky",r"mizuno",r"\bkuhl\b",r"nobull",r"wolfpak",r"gym ?reaper",r"\blskd\b",r"\bketl\b",r"undeniable",r"rvca",r"\bmitre\b",r"popflex",r"equinox",r"helly hansen",r"hyperice",r"born primitive",r"southern athletica",r"lift threadz",r"paragon running",r"\bcentric\b",r"greyson",r"ryker gym",r"\bzivi\b",r"raide shorts",r"minted ?ny|minted new york|minted running",r"randomevent",r"rabbit clothing",r"legends clothing",r"muscle republic",r"\byiwo\b",r"freetrain",r"wolaco",r"herman koumori",r"aesene",r"kill crew",r"aswemove",r"feisupao",r"mover activewear",r"everlast",r"fgm04",r"leorever",r"coyote clothing",r"\bgbrs\b",r"spyder",r"malbon",r"fear of god",r"cove t shirt",r"gg short",r"meat shorts",r"raidlight",r"westslope",r"angry pablo",r"headlands",r"deadlast",r"\bvhts\b",r"anonymous morning",r"patencials",r"is street outwears? legit|street outerwears")
A_ELIM=[(R(r"bib shorts",r"cycling",r"chamois",r"padded bik",r"bike shorts",r"\bbibs\b",r"\bmtb\b"),"Cycling/bike product (padded or bib shorts) outside ASRV's training and streetwear range"),
        (R(r"texas a&m"),"Collegiate-merchandise query with no match to ASRV products"),
        (R(r"skull caps for"),"Informational question with no purchase intent"),
        (R(r"wrestling shorts",r"fishing leggings"),"Niche sport/activity apparel outside ASRV's product range")]
A_CONS=[(R(r"sports bra",r"bra top",r"longline",r"sheer gym top",r"women'?s running briefs"),"Women's sports-bra/undergarment style term; adjacent, confirm it fits ASRV's women's range"),
        (R(r"tennis",r"pickleball",r"\bgolf\b",r"muay thai",r"vale tudo",r"fighting shorts",r"triathlon",r"tri shorts",r"crossfit apparel",r"hot yoga",r"pilates",r"upf 50"),"Adjacent sport category outside ASRV's core training/streetwear positioning"),
        (R(r"near me"),"'Near me' phrasing signals local-retail intent for a DTC brand"),
        (R(r"what do guys wear",r"harry styles"),"Style-research/celebrity query; not clear purchase intent"),
        (R(r"[^\x00-\x7f]"),"Non-English query; confirm language/market targeting"),
        (R(r"^camouflage$",r"^on$"),"Single generic word; intent unclear")]
A_BRANDMOD=R(r"discount",r"coupon",r"promo",r"\bcodes?\b",r"free shipping",r"stand for",r"reddit",r"review",r"legit")
def cl_asrv(x):
    t=x['term'];conv=x['conv']
    if A_COMP.search(t): return ('COMPETITOR',"Contains a non-ASRV brand name; may be intentional conquesting, so review only")
    if A_BRAND.search(t):
        if conv==0 and A_BRANDMOD.search(t): return ('CONSIDER',"ASRV brand + coupon/review-style modifier; brand term, so check match type or coupon-site leakage rather than negating")
        return None
    if conv==0:
        for rx,rs in A_ELIM:
            if rx.search(t): return ('ELIMINATE',rs)
    if conv==0:
        for rx,rs in A_CONS:
            if rx.search(t): return ('CONSIDER',rs)
        if x['spend']>=50: return band(x,50)
    return None

# ---------------- GRIFFIN ----------------
G_COMP={"prime direct capital","the loan guys","mgic self employed","lending one dscr","emd lender","dunmor lending","acra lending dscr","direct private capital group","nation capital","r e lending","velocity fast 50","easy street lending","pillar private lending","landlord loan enhancements the mortgage works","trilith funding"}
G_ELIM={"bank statement car loan":"Auto-loan query; Griffin does not offer car loans",
 "car loan with bank statements":"Auto-loan query; Griffin does not offer car loans",
 "cmbs financing":"Commercial real-estate financing; Griffin does not do commercial loans",
 "hud multifamily loans":"Government commercial multifamily financing; Griffin does not do commercial loans",
 "usda loans for investment property":"USDA loans are not offered and are not available for investment property",
 "dscr sba loan":"SBA loan query; not a Griffin product",
 "can i use a sba loan to buy real estate":"SBA loan query; not a Griffin product",
 "business loans for self employed":"General business-loan intent; Griffin does not do commercial loans",
 "instant cash for self employed":"Cash-advance intent, not mortgage lending",
 "cash advances for self employed":"Cash-advance intent, not mortgage lending",
 "cash advance self employed":"Cash-advance intent, not mortgage lending",
 "quick loan for self employed":"Short-term personal-loan intent, not mortgage lending",
 "self employed loans guaranteed approval":"Guaranteed-approval personal-loan intent, not mortgage lending",
 "example of mortgage statement":"Existing-borrower statement lookup; no loan-shopping intent",
 "mortgage statement":"Existing-borrower statement lookup; no loan-shopping intent",
 "family buy to let mortgage":"UK 'buy to let' phrasing; wrong geography"}
G_CONS=[(R(r"\bhow to\b",r"\bwhat is\b",r"\bwhat does\b",r"^explain",r"\bexplain\b",r"how does",r"how dscr",r"calculator",r"calculate",r"^what are"),"Informational/calculator phrasing; may be research-stage rather than buyer intent"),
 (R(r"\b(dhcr|drcs?|dcra|dcrs|dsrc|dser|dsr|dsc|dcsr)\b"),"Possible misspelling or unrelated acronym of 'DSCR'; verify before acting"),
 (R(r"\bfha\b"),"FHA loans are not in Griffin's listed product mix; confirm fit"),
 (R(r"manufactured home"),"Manufactured-home financing; confirm Griffin underwrites this property type"),
 (R(r"distressed property"),"Distressed-property/bridge financing; adjacent to Griffin's products"),
 (R(r"vacation property"),"Vacation-property financing; confirm program fit"),
 (R(r"build to rent",r"fix and rent"),"Construction/rehab-style investor financing; adjacent to DSCR"),
 (R(r"multi unit",r"multifamily"),"Multifamily/multi-unit; 5+ units may fall under commercial lending, which Griffin does not do"),
 (R(r"\bbtl\b"),"'BTL' is UK buy-to-let terminology; possible wrong geography"),
 (R(r"prestamo"),"Spanish-language query; confirm language targeting"),
 (R(r"cpa letter"),"Documentation/how-to style query rather than lender search"),
 (R(r"business (bank statement )?loans",r"bank statement business",r"business bank statement"),"Business-loan phrasing; Griffin does not do commercial loans"),
 (R(r"loan places"),"Consumer 'loan places' phrasing; likely personal-loan intent"),
 (R(r"griffin funding reviews",r"griffin funding reviews complaints"),"Brand + review modifier; brand term, so review match type rather than negating")]
G_NOINC=re.compile(r"\bloans?\b",re.I)
G_NOINC_MOD=re.compile(r"proof of income|without income|no income|income verification|pay ?stub|job verification|income check|require income|verify income|without proof|no proof|without job|need proof",re.I)
G_NOINC_EXC=re.compile(r"mortgage|home|house|real estate|property|rental|investment|dscr|bank statement|self employed|1099|heloc|refinanc|independent contractor|freelanc",re.I)
def cl_griffin(x):
    t=x['term'];conv=x['conv']
    if t in G_COMP: return ('COMPETITOR',"Appears to be another lender's brand name; may be intentional conquesting, so review only")
    if conv==0 and t in G_ELIM: return ('ELIMINATE',G_ELIM[t])
    if conv==0:
        for rx,rs in G_CONS:
            if rx.search(t): return ('CONSIDER',rs)
        if G_NOINC.search(t) and G_NOINC_MOD.search(t) and not G_NOINC_EXC.search(t):
            return ('CONSIDER',"Generic 'no income proof' loan query with no mortgage wording; may be personal-loan intent")
    return None

# ---------------- SPOTLIGHT ----------------
SP_COMP=R(r"tix ?4 ?tonight",r"tix4vegas",r"tix for tonight",r"tickets 4 tonight",r"tix4 com",r"tix 4 tonight",r"\bvegas ?com\b",r"las ?vegas ?com\b",r"lasvegas com",r"groupon",r"costco",r"viator",r"getyourguide",r"ticketmaster",r"stubhub",r"seatgeek",r"vivid ?seats",r"tripadvisor",r"expedia",r"goldstar")
SP_ELIM=[(R(r"londres",r"londra",r"london",r"伦敦"),"Show query for London, outside Las Vegas ticketing"),
         (R(r"beto carrero"),"Brazilian theme-park query unrelated to Las Vegas shows"),
         (R(r"magic mike new york"),"New York show query, outside Las Vegas ticketing"),
         (R(r"\bjobs?\b",r"careers?",r"hiring",r"salary",r"audition"),"Job-seeker query with no ticket-buying intent"),
         (R(r"restaurant"),"Restaurant query; Spotlight sells show/attraction tickets, not dining"),
         (R(r"what is las vegas strip",r"max flight simulator"),"Informational/unrelated attraction query with no ticket intent")]
SP_HOTEL={"mgm grand","excalibur hotel","excalibur","rio las vegas","flamingo las vegas","horseshoe las vegas","treasure island las vegas","caesars palace las vegas","luxor las vegas","harrah's las vegas","harrahs las vegas","stratosphere","planet hollywood","planet hollywood las vegas","planet hollywood vegas","planet hollywood in vegas","mgm grand las vegas","caesar shows","luxor hotel las vegas","riu hotel las vegas","pyramid hotel vegas","flamingo hotel las vegas","rio hotel las vegas","luxor hotel","stratosphere hotel","caesars palace hotel","mandalay bay hotel","hotel mandalay bay las vegas","excalibur vegas","las vegas excalibur","paris las vegas","new york new york las vegas","excalibur las vegas","the palazzo","excalibur hotel las vegas","mgm luxor hotel"}
SP_CONS=[(R(r"\bfree\b"),"'Free' modifier; may be seeking no-cost activities rather than paid tickets"),
         (R(r"near me"),"'Near me' phrasing with no Las Vegas anchor"),
         (R(r"strip club",r"nightclub",r"day club",r"pool party",r"adult club",r"male strip club"),"Nightlife/club query; adjacent to Spotlight's show inventory"),
         (R(r"^what is\b",r"\bwhat is\b"),"'What is' informational phrasing")]
def cl_spot(x):
    t=x['term'];conv=x['conv']
    if SP_COMP.search(t): return ('COMPETITOR',"Contains a competing ticket seller/marketplace name; may be intentional conquesting, so review only")
    if conv==0:
        for rx,rs in SP_ELIM:
            if rx.search(t): return ('ELIMINATE',rs)
    if conv==0:
        if t in SP_HOTEL: return ('CONSIDER',"Hotel/venue navigational query; may be hotel-booking rather than show-ticket intent")
        for rx,rs in SP_CONS:
            if rx.search(t): return ('CONSIDER',rs)
        if x['spend']>=50: return band(x,50)
    return None

# ---------------- LEARN ----------------
L_COMP=set("""circle city aba|butterfly effects aba|cortica braintree|majestic care aba crown point|kind behavioral health greensboro|ally behavior centers fairfax|ally aba therapy|ally behavior centers|ally behavior center|ally behavior centers springfield|ally behavior|centria autism indianapolis|centria autism boston ma|centria autism grand rapids|above and beyond aba therapy|above and beyond therapy|acorn health|proud moments aba silver spring|proud moments aba|achievements aba therapy|behavior learning tree|brighter strides aba|brighter strides aba therapy|autism first chantilly|autism f i r s t|soar autism center rockville|soar autism center silver spring md|pbs aba therapy|lurie center lexington ma|hope bridge autism center|hope behavioral services logansport|alpaca health|next step behavioral services|eden aba therapy|mindsparks therapy|cabi core auburn ma|verbal beginnings rockville|verbal beginnings columbia|verbal beginnings frederick|aligned vision behavior|total care therapy aba|golden steps aba|golden step aba|golden steps aba phone number|oos management corp aba|kaleidoscope aba|aba centers of america worcester|aba centers of america peabody|aba centers of america foxboro ma|abs kids phone number|early autism solutions dc|ascend aba bronx|mass pediatric aba llc|unc teacch raleigh|autism learning partners|roots to blooms aba|aba today reston|autism care partners worcester|afl autism services adapt for life|meaningful outcomes aba|kids first aba therapy|children milestones aba|little star west lafayette indiana|autism spectrum therapies ast|https user teampbs com process|https myasdf org our programs""".split('|'))
L_ELIM={"aba training":"Professional training/certification query; not a family seeking services",
"aba therapy training online free":"Professional training query with 'free' modifier; not a family seeking services",
"aba certification massachusetts":"Professional certification query; not a family seeking services",
"school rbt":"Technician job/training query; not a family seeking services",
"autism partnership foundation 40 hour training":"Professional training query; not a family seeking services",
"free programs for autistic adults":"Free adult-program query; outside LEARN's early-intervention focus on children and young adults",
"housing help for autistic children":"Housing-assistance query; outside the services LEARN provides",
"l1 79 treatment":"Unclear or garbled query with no coherent intent",
"red light therapy brain autism":"Alternative-treatment query unrelated to ABA services"}
L_CONS={}
for k in ["autism care today grant","act grants autism","grants for autism","autism care today grant 2026 application","autism care today grant 2026","autism waiver maryland"]:
    L_CONS[k]="Funding/grant research; families may need financial help, but intent is not a provider search"
L_CONS["free services for autistic children"]="'Free' modifier; families may be looking for no-cost services"
for k in ["autism behavior therapy for adults","aba therapy for adults near me"]:
    L_CONS[k]="Adult-services intent; outside LEARN's early-intervention focus on children (verify local availability)"
for k in ["school for autism near me","schools for autistic children","preschool for autistic children near me","after school programs for autistic children","respite care for children with autism","autism social groups","peers social skills group near me","play groups for autistic kids","life skills for autistic teens","iep for autism","autism occupational therapy near me","autism speech therapy"]:
    L_CONS[k]="Adjacent support service (education, respite, social or related therapy) rather than ABA therapy"
L_CONS["aba therapy for down syndrome"]="Adjacent condition; unclear whether LEARN's program fits"
L_CONS["how to get child tested for adhd and autism"]="Combined ADHD/autism evaluation query; partly outside LEARN's diagnostic focus"
L_CONS["autistic level 1 support needs"]="Informational query about support needs; research-stage intent"
L_CONS["explain how aba therapy works for autism and what techniques are used"]="Informational 'explain' phrasing; research-stage intent"
L_CONS["what is aba therapy"]="'What is' informational phrasing; research-stage intent"
L_CONS["bcba near me"]="Unclear intent: could be a family seeking a clinician or a professional seeking work"
L_CONS["aba terapia para autismo"]="Spanish-language query; confirm language targeting and availability"
L_CONS["bryces journey"]="Unclear intent (possibly a family-story or brand query) with notable spend"
L_CONS["help is in your hands autism"]="Unclear intent; may refer to another organization's resource"
def cl_learn(x):
    t=x['term'];conv=x['conv']
    if t in L_COMP: return ('COMPETITOR',"Appears to reference a non-LEARN provider; may be intentional conquesting, so review only")
    if conv==0 and t in L_ELIM: return ('ELIMINATE',L_ELIM[t])
    if conv==0 and t in L_CONS: return ('CONSIDER',L_CONS[t])
    if conv==0 and x['spend']>=200: return band(x,200)
    return None

# ---------------- ECOPEACO ----------------
E_BRAND=R(r"eco ?pea")
E_COMP=R(r"aleva",r"\bhonest\b",r"pharmaprix",r"royale",r"cotocouche",r"kirkland",r"petits culott",r"coterie",r"love diaper brand",r"well ca",r"\bbambo\b",r"hello bello",r"snuggle bugz",r"kudos")
E_CONS=[(R(r"reusable",r"washable",r"\bcloth\b",r"booster pads"),"Reusable/cloth-diaper or add-on intent; adjacent to EcoPeaCo's disposable range"),
        (R(r"pañales"),"Spanish-language query; confirm language targeting"),
        (R(r"how does",r"reddit"),"Informational/forum phrasing; research stage")]
def cl_eco(x):
    t=x['term'];conv=x['conv']
    if E_COMP.search(t): return ('COMPETITOR',"Contains a non-EcoPeaCo brand or retailer name; may be intentional conquesting, so review only")
    if conv==0 and E_BRAND.search(t):
        if re.search(r"review|where are|made",t): return ('CONSIDER',"EcoPeaCo brand + review/informational modifier; brand term, so check match type rather than negating")
        if x['spend']>=50: return ('CONSIDER',f"Brand term with no conversions on ${x['spend']:,.0f} spend; high for brand traffic, so check match type, bids and ad-group placement")
        return None
    if conv==0:
        if re.search(r"diapers that start with c",t): return ('ELIMINATE',"Trivia-style query with no purchase intent")
        if re.search(r"uber eats gift card",t): return ('ELIMINATE',"Gift card for an unrelated brand; not a diaper-purchase query")
        for rx,rs in E_CONS:
            if rx.search(t): return ('CONSIDER',rs)
        if x['spend']>=50: return band(x,50)
    return None

# ---------------- WE BUY EXOTICS ----------------
W_BRAND=R(r"we ?buy ?exotic",r"webuyexotics")
W_COMP=R(r"i ?lusso|illuso|illusso",r"j ?(&|and) ?s auto",r"jsautohaus",r"exotic car trader",r"sellyoursportscar")
def cl_wbe(x):
    t=x['term'];conv=x['conv']
    if W_COMP.search(t): return ('COMPETITOR',"Matches a competing dealer/marketplace (e.g. iLusso, J&S Autohaus); sits in the conquest campaign, so review only")
    if W_BRAND.search(t):
        if conv==0 and re.search(r"review|\d{3,}",t): return ('CONSIDER',"Brand + review/street-address lookup; brand term, so review match type rather than negating")
        if conv==0 and x['spend']>=50: return ('CONSIDER',f"Brand term with no conversions on ${x['spend']:,.0f} spend; check match type and bids")
        return None
    if conv==0:
        if re.search(r"import price",t): return ('ELIMINATE',"Buyer-side import pricing research; WBE buys vehicles from sellers")
        if re.search(r"dealer|broker|agency",t): return ('CONSIDER',"Dealer/broker-seeking phrasing; ambiguous between sellers and shoppers")
        if re.search(r"consignment",t): return ('CONSIDER',"Consignment intent; WBE buys outright, so intent may differ")
        if re.search(r"used mk4 supra|supra mk4 price",t): return ('CONSIDER',"Buyer-side price research for a specific model")
        if re.search(r"worth|exotic car market|aston martin car",t): return ('CONSIDER',"Valuation/market-research query; may not be ready-to-sell intent")
        if re.search(r"^how to sell",t): return ('CONSIDER',"'How to' informational phrasing; research stage")
        if re.search(r"sell my audi",t): return ('CONSIDER',"Non-exotic marque; confirm WBE buys this vehicle type")
        if re.search(r"exsotic",t): return ('CONSIDER',"Misspelled generic exotic-car query; intent unclear")
        if x['spend']>=50: return band(x,50)
    return None

# ---------------- GOGATHER ----------------
def cl_gg(x):
    t=x['term'];conv=x['conv']
    if re.search(r"freeman",t): return ('COMPETITOR',"Names a known event-services provider (Freeman); may be intentional conquesting, so review only")
    if conv==0:
        if re.search(r"venue|event space|event centers?|rental space",t): return ('CONSIDER',"Venue-rental intent; confirm GoGather offers venue sourcing rather than planning/destination services")
        if re.search(r"team building|entertainment",t): return ('CONSIDER',"Adjacent corporate activity/entertainment query; confirm service fit")
        if x['spend']>=50: return band(x,50)
    return None

# ---------------- RCM ----------------
RC_ELIM={"https tcs adp com comsys":"Payroll-portal/URL fragment with no relationship to tax-credit services",
"residential energy credit":"Consumer residential credit query; RCM serves commercial credit transfers",
"residential clean energy credit":"Consumer residential credit query; RCM serves commercial credit transfers",
"u s small business administration washington metropolitan district":"Government agency navigational query with no purchase intent",
"tennessee llc taxes":"General state LLC tax query unrelated to energy credits",
"wa business license":"Business-licensing query unrelated to energy credits"}
RC_COMP={"ion energy broker":"Names another energy-credit broker (ION Energy); may be intentional, so review only",
         "world tax services":"Names another tax-services firm; may be intentional, so review only"}
RC_CONS={"solar tax credit":"Solar credit phrasing; could be a business with credits to sell or a consumer",
"solar itc credit":"Solar ITC phrasing; could be a business with credits to sell or a consumer",
"tax credit for solar panels":"Solar-panel credit phrasing; leans consumer, but unclear",
"trade america":"Unclear intent; possibly another company or trade program",
"oseries":"Unclear intent; possibly a product or competitor name",
"renewable energy law":"Legal-information query; research stage",
"partnership":"Single generic word; intent unclear",
"company in usa":"Vague generic query; intent unclear"}
def cl_rcm(x):
    t=x['term'];conv=x['conv']
    if t in RC_COMP: return ('COMPETITOR',RC_COMP[t])
    if conv==0 and t in RC_ELIM: return ('ELIMINATE',RC_ELIM[t])
    if conv==0 and t in RC_CONS: return ('CONSIDER',RC_CONS[t])
    if conv==0 and x['spend']>=50: return band(x,50)
    return None

# ---------------- small full-tier ----------------
def cl_eurolog(x):
    t=x['term']
    if re.search(r"conex|container foam insulation",t): return ('CONSIDER',"Container-insulation phrasing; often DIY/storage (container-home crossover) rather than B2B export buyers")
    if re.search(r"heating blanket",t): return ('CONSIDER',"Heating-blanket phrasing; adjacent to thermal stabilization products")
    return None
SUN={"site shades":("CONSIDER","Unclear intent; may be a generic shade query or a Sunrise SHADE variant"),
"tree netting for climbing":("CONSIDER","Recreational 'climbing' phrasing; Sunrise netting targets architectural applications"),
"hammock net":("CONSIDER","Generic hammock phrasing; may be consumer-hammock intent rather than overwater/architectural netting"),
"besenzoni parts":("COMPETITOR","Names a marine-accessory manufacturer (Besenzoni); may be intentional, so review only"),
"vetus catalogue":("COMPETITOR","Names a marine-parts brand (Vetus); catalogue-seeking, so review only")}
def cl_sun(x): return SUN.get(x['term'])
TRI={"outdoor powerlifting gym":("CONSIDER","Gym-style phrasing may indicate a private/home gym buyer rather than a government buyer"),
"elliptical machine for outdoors":("CONSIDER","Consumer-style product phrasing; TriActive sells commercial park and correctional equipment"),
"outdoor workout machine":("CONSIDER","Consumer-style product phrasing; TriActive sells commercial park and correctional equipment")}
def cl_tri(x): return TRI.get(x['term'])
def cl_none(x): return None

# ---------------- run ----------------
LEARN_CODES={'WL':'SMM_NBS_INDIANA_WestLafayette','FI':'SMM_NBS_INDIANA_Fishers','FR':'SMM_NBS_EMASS_FallRiver','QU':'SMM_NBS_EMASS_Quincy','CP':'SMM_NBS_INDIANA_CrownPoint','CA':'SMM_NBS_NCAROL_Cary','RE':'SMM_NBS_VA_Reston','GW':'SMM_NBS_INDIANA_Greenwood','CH':'SMM_NBS_VA_Chantilly','AL':'SMM_NBS_VA_Alexandria','AS':'SMM_NBS_VA_Ashburn','TY':'SMM_NBS_VA_Tysons_Corner','GE':'SMM_NBS_MD_Germantown','WN':'SMM_NBS_EMASS_WorcesterN','FM':'SMM_NBS_EMASS_Framingham','RO':'SMM_NBS_EMASS_Roxbury','ST':'SMM_NBS_EMASS_Sturbridge','LI':'SMM_NBS_EMASS_Littleton','SF':'SMM_Sparks_Fairfax VA','SM':'SMM_Sparks_Manassas VA','SS':'SMM_Sparks_Springfield VA','GR':'SMM_Total Spectrum_Grand Rapids MI','KA':'SMM_Total Spectrum_Kalamazoo MI','LL':'SMM_Little Leaves_Columbia MD'}
ECA={'BC':'BFD_Brand Core_Exact_CA','SN':'BFD_Shopping_NB_CA','SB':'BFD_Shopping_Brand_CA','NC':'BFD_NB Core_Exact_CA','NG':'BFD_NB Gifting_Exact_CA'}
EUS={'BC':'BFD_Brand Core_Exact_US','SB':'BFD_Shopping_Brand_US','NC':'BFD_NB Core_Exact_US','BW':'BFD_NB_Shopping_Bamboo'}
WB={'BR':'SMM_BS_We_Buy_Exotics','EC':'SMM_NBS_Exotic_Cars','CO':'SMM_NBS_Competitors','LT':'SMM_NBS_Exotic_Cars_Low tCPA'}
GGC={'AE':'BFD_GoGather_NB_Search_All Events','DM':'BFD_GoGather_NB_Search_All Events','BR':'BFD_GoGather_B_Search_Brand','DS':'BFD_GoGather_NB_Search_Destination','DD':'BFD_GoGather_NB_Search_Destination','DV':'BFD_GoGather_NB_Search_Destination','DF':'BFD_GoGather_NB_Search_Destination','DN':'BFD_GoGather_NB_Search_Destination','DO':'BFD_GoGather_NB_Search_Destination'}
RCC={'PE':'BFD_RCM_Google_NB_Search_Pros_Energy Credit','IS':'BFD_RCM_Google_NB_Search_Impression_Share','BR':'BFD_RCM_Google_B_Search_Pros_Energy Credit'}
EUC={'NB':'BFD_NB_Products'}
SUC={'T2':'SMM_NBS_Catamaran_T2','MH':'multihull nets sales'}
TRC={'NB':'SMM_NBS_Primary_KWs','BS':'SMM_BS_TriActiveUSA'}
RPC={'RP':'RipCo Products'}

ACCTS=[
 ('ASRV',load_json('asrv.json'),cl_asrv),
 ('Griffin',load_json('griffin.json'),cl_griffin),
 ('Velo',None,None),
 ('Sunrise',agg_raw(parse_small('SUNRISE'),SUC),cl_sun),
 ('TriActive',agg_raw(parse_small('TRIACTIVE'),TRC),cl_tri),
 ('LEARN',agg_raw(S+'learn_raw.txt',LEARN_CODES),cl_learn),
 ('Eurolog',agg_raw(parse_small('EUROLOG'),EUC),cl_eurolog),
 ('Spotlight',load_json('spot.json'),cl_spot),
 ('CQC',None,None),
 ('EcoPeaCo CA',agg_raw(S+'eco_ca_raw.txt',ECA),cl_eco),
 ('EcoPeaCo US',agg_raw(S+'eco_us_raw.txt',EUS),cl_eco),
 ('GoGather',agg_raw(S+'gg_raw.txt',GGC),cl_gg),
 ('RCM',agg_raw(S+'rcm_raw.txt',RCC),cl_rcm),
 ('Ripco',agg_raw(parse_small('RIPCO'),RPC),cl_none),
 ('We Buy',agg_raw(S+'wbe_raw.txt',WB),cl_wbe),
]
out={}
for name,rows,fn in ACCTS:
    if rows is None: out[name]=None; continue
    fl=[]
    for x in rows:
        r=fn(x)
        if r: fl.append(dict(cat=r[0],term=x['term'],camps=x['camps'],imp=x['imp'],clicks=x['clicks'],spend=round(x['spend'],2),conv=round(x['conv'],2),val=round(x['val'],2),reason=r[1]))
    order={'ELIMINATE':0,'CONSIDER':1,'COMPETITOR':2}
    fl.sort(key=lambda f:(order[f['cat']],-f['spend']))
    out[name]=dict(total=len(rows),flagged=fl,total_spend=round(sum(x['spend'] for x in rows),2))
    c=collections.Counter(f['cat'] for f in fl)
    print(f"{name:12s} terms={len(rows):5d} E={c['ELIMINATE']:3d} C={c['CONSIDER']:3d} X={c['COMPETITOR']:3d}")
json.dump(out,open(S+'results.json','w'))
