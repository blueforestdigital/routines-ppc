#!/usr/bin/env python3
"""Weekly Negative Keyword Analysis — 2026-05-12"""
import json, os, re
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

DATA_DIR = "/root/.claude/projects/-home-user-routines-ppc/41541e4f-07c4-4a02-a584-58b9a0e1e722/tool-results"
REPORT_DATE = "2026-05-12"
DATE_RANGE = "May 5–11, 2026"

# ── inline data for 6 newly-fetched accounts ────────────────────────────────
NEW_DATA = {
"Eurolog": [{"search_term":"container liners","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":1,"spend":10.82,"conversions":0,"conversion_value":0},{"search_term":"desiccant bags for shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":16,"clicks":2,"spend":20.18,"conversions":0,"conversion_value":0},{"search_term":"shipping container moisture control","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":1,"spend":4.9439,"conversions":0,"conversion_value":0},{"search_term":"insulated pallet blanket","campaign":"BFD_NB_AllCampaigns","impressions":11,"clicks":1,"spend":5.9637,"conversions":0,"conversion_value":0},{"search_term":"desiccant bags for containers","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":1,"spend":1.95,"conversions":0,"conversion_value":0},{"search_term":"container liner shipping","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":2,"spend":30.32,"conversions":2,"conversion_value":0},{"search_term":"celluliner","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":1,"spend":18.34,"conversions":0,"conversion_value":0},{"search_term":"shipping container foil liner","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":1,"spend":2.6086,"conversions":0,"conversion_value":0},{"search_term":"bulk container liner","campaign":"BFD_NB_AllCampaigns","impressions":5,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container liner","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"sea bulk container liners","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container desiccant bags","campaign":"BFD_NB_AllCampaigns","impressions":5,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container desiccants","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"dry bags for shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"how to keep a conex from sweating","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"shipping container desiccant","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"foil box liner","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"freezer blanket for food transport","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"insulated pallet covers","campaign":"BFD_NB_AllCampaigns","impressions":7,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pallet thermal blankets","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal blanket for shipping","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal covers","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal liner container","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"vineliner","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"bulk container liners","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"absorgel desiccant","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"absorgel hanging","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container dri","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container moisture control","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"dehumidifier for shipping container","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"desiccant bags for containers","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":1,"spend":1.95,"conversions":0,"conversion_value":0},{"search_term":"how to stop shipping container from sweating","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"cargo quilts","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"climaliner","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"insulation panels for shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pallet insulation covers","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"silverskin thermal blanket","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal blanket wrap","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"bulk sea container","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"liner bags for containers","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"sea bulk container","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"absortech desiccant","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"cargo desiccant bag","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container desiccant","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seadry desiccant","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"shipping container desiccants","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container thermal liner","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pallet blankets","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal blanket for food transportation","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal blanket for frozen food","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal box panels","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal container liner","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal moving blankets","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"clariant container dri ii","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"condensation in shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container dri ii desiccants cargo bag","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"moisture in shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"cargo blankets","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"dock blanket","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"heated pallet blanket","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"industrial insulation blankets","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"insulated cargo blanket","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"insulation for conex container","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ipc liners","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pallet cover insulated","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"bulk shipping containers","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container dehumidifier","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container dryer humidity absorber","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"dry pole desiccant","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"topcod desiccant","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"container blanket","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"insulated pallet cover","campaign":"BFD_NB_AllCampaigns","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pet fibre liners for insulated packaging","campaign":"BFD_NB_AllCampaigns","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal box liner","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal liners","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"thermal pallet covers","campaign":"BFD_NB_AllCampaigns","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"xr31225gr","campaign":"BFD_NB_AllCampaigns","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0}],
"Sunrise": [{"search_term":"multihull nets","campaign":"multihull nets sales","impressions":4,"clicks":3,"spend":2.7531,"conversions":5.5727,"conversion_value":3183.816},{"search_term":"outdoor loft netting","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":4.0347,"conversions":0,"conversion_value":0},{"search_term":"trampoline catamaran","campaign":"multihull nets sales","impressions":4,"clicks":2,"spend":1.0584,"conversions":0,"conversion_value":0},{"search_term":"loft nets","campaign":"multihull nets sales","impressions":8,"clicks":2,"spend":0.5941,"conversions":0,"conversion_value":0},{"search_term":"trimaran trampoline","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":2.66,"conversions":0,"conversion_value":0},{"search_term":"lagoon 50","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":2.17,"conversions":1,"conversion_value":0.75},{"search_term":"treehousesupplies","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":1.3,"conversions":0,"conversion_value":0},{"search_term":"catamaran netting","campaign":"SMM_NBS_Catamaran","impressions":3,"clicks":1,"spend":11.53,"conversions":0,"conversion_value":0},{"search_term":"multihull net","campaign":"SMM_NBS_Catamaran","impressions":3,"clicks":1,"spend":0.36,"conversions":0,"conversion_value":0},{"search_term":"sailboat trampoline","campaign":"SMM_NBS_Catamaran","impressions":1,"clicks":1,"spend":4.4389,"conversions":0,"conversion_value":0},{"search_term":"catamaran trampoline for sale","campaign":"SMM_NBS_Catamaran_T2","impressions":1,"clicks":1,"spend":3.9546,"conversions":0,"conversion_value":0},{"search_term":"catamaran tramp","campaign":"SMM_NBS_Catamaran_T2","impressions":2,"clicks":1,"spend":3.37,"conversions":0,"conversion_value":0},{"search_term":"trampoline for catamaran","campaign":"SMM_NBS_Catamaran_T2","impressions":2,"clicks":1,"spend":2.16,"conversions":1,"conversion_value":0.75},{"search_term":"hobie cat trampoline","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":0.38,"conversions":0,"conversion_value":0},{"search_term":"trampolin für katamaran","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":1.75,"conversions":0,"conversion_value":0},{"search_term":"sailboat hammock","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":0.23,"conversions":0,"conversion_value":0},{"search_term":"dock nets","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":0.36,"conversions":0,"conversion_value":0},{"search_term":"catamaran trampoline material","campaign":"multihull nets sales","impressions":1,"clicks":1,"spend":1.23,"conversions":0,"conversion_value":0},{"search_term":"multihulls net","campaign":"SMM_NBS_Catamaran_T2","impressions":2,"clicks":1,"spend":0.26,"conversions":0,"conversion_value":0},{"search_term":"multihull nets","campaign":"SMM_NBS_Catamaran","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"multihull nets","campaign":"SMM_NBS_Catamaran_T2","impressions":4,"clicks":1,"spend":0.01,"conversions":0,"conversion_value":0},{"search_term":"catamaran trampoline","campaign":"SMM_NBS_Catamaran","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran netting","campaign":"multihull nets sales","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran trampoline netting","campaign":"multihull nets sales","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"dock net","campaign":"multihull nets sales","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"loft net","campaign":"multihull nets sales","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"loftnet","campaign":"multihull nets sales","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"trampolines for sale","campaign":"multihull nets sales","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"boat for sale","campaign":"multihull nets sales","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran","campaign":"multihull nets sales","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran net","campaign":"multihull nets sales","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"hobie trampoline","campaign":"multihull nets sales","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"large hammock net","campaign":"multihull nets sales","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"tree net","campaign":"multihull nets sales","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"trampoline catamaran","campaign":"SMM_NBS_Catamaran","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran net","campaign":"SMM_NBS_Catamaran","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"catamaran trampoline","campaign":"SMM_NBS_Catamaran_T2","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0}],
"TriActive": [{"search_term":"calisthenic park equipment","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":1.86,"conversions":0,"conversion_value":0},{"search_term":"used outdoor exercise equipment","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":2.57,"conversions":0,"conversion_value":0},{"search_term":"gym equipment for outdoor use","campaign":"SMM_NBS_Primary_KWs","impressions":4,"clicks":1,"spend":0.4851,"conversions":0,"conversion_value":0},{"search_term":"public park exercise equipment","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":12.3394,"conversions":0,"conversion_value":0},{"search_term":"triactive america","campaign":"SMM_NBS_Primary_KWs","impressions":3,"clicks":1,"spend":9.4501,"conversions":0,"conversion_value":0},{"search_term":"best weatherproof outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":1,"spend":10.3219,"conversions":0,"conversion_value":0},{"search_term":"outdoor park gym equipment","campaign":"SMM_NBS_Primary_KWs","impressions":3,"clicks":2,"spend":6.37,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","impressions":41,"clicks":3,"spend":5.35,"conversions":0,"conversion_value":0},{"search_term":"exercise stations park","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":3.96,"conversions":0,"conversion_value":0},{"search_term":"outdoor dip bars","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":1.0264,"conversions":0,"conversion_value":0},{"search_term":"park elliptical","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":1,"spend":5.6976,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym","campaign":"SMM_NBS_Primary_KWs","impressions":22,"clicks":1,"spend":1.026,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym equipment for parks","campaign":"SMM_NBS_Primary_KWs","impressions":7,"clicks":1,"spend":2.14,"conversions":0,"conversion_value":0},{"search_term":"weatherproof workout equipment","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":1,"spend":1.0039,"conversions":0,"conversion_value":0},{"search_term":"triactive","campaign":"SMM_BS_TriActiveUSA","impressions":20,"clicks":3,"spend":4.37,"conversions":0,"conversion_value":0},{"search_term":"triactive usa","campaign":"SMM_BS_TriActiveUSA","impressions":37,"clicks":8,"spend":7.5207,"conversions":1,"conversion_value":0},{"search_term":"outdoor bodyweight workout equipment","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":1,"spend":1.21,"conversions":0,"conversion_value":0},{"search_term":"outdoor park workout equipment","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":1,"spend":0.53,"conversions":0,"conversion_value":0},{"search_term":"outdoor dip station","campaign":"SMM_NBS_Primary_KWs","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym equipment for adults","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"park elliptical","campaign":"SMM_NBS_Primary_KWs","impressions":2,"clicks":1,"spend":5.6976,"conversions":0,"conversion_value":0},{"search_term":"diy outdoor exercise stations","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"kids outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison yard gym","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"exercise equipment for parks","campaign":"SMM_NBS_Primary_KWs","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0}],
"Velo": [{"search_term":"go city wide","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":1,"spend":8.86,"conversions":0,"conversion_value":0},{"search_term":"commercial office cleaning","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":1,"spend":7.22,"conversions":0,"conversion_value":0},{"search_term":"restaurant cleaning","campaign":"BFD_Velo_Google_NB_Search_Pros_T3","impressions":3,"clicks":1,"spend":10.2748,"conversions":0,"conversion_value":0},{"search_term":"janitorial franchise","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":2,"clicks":1,"spend":14.93,"conversions":0,"conversion_value":0},{"search_term":"pro serve cleaning","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":2,"clicks":1,"spend":24.06,"conversions":0,"conversion_value":0},{"search_term":"building cleaning services near me","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":1,"spend":19.57,"conversions":0,"conversion_value":0},{"search_term":"construction cleaning services","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":1,"spend":14.4702,"conversions":0,"conversion_value":0},{"search_term":"office cleaning packages","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":1,"spend":22.41,"conversions":0,"conversion_value":0},{"search_term":"janitorial franchise","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":2,"clicks":1,"spend":14.93,"conversions":0,"conversion_value":0},{"search_term":"pudu floor scrubber","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"kaivac kv1750","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"floor cleaning industrial machines","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"best commercial pressure washer","campaign":"BFD_Velo_Google_NB_Search_Pros","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"power washing phoenix","campaign":"BFD_Velo_Google_NB_Search_Pros_T3","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pressure washing phoenix","campaign":"BFD_Velo_Google_NB_Search_Pros_T3","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0}],
"CQC": [{"search_term":"golf cart seat covers","campaign":"CQC_NBS","impressions":191,"clicks":15,"spend":34.5736,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for yamaha","campaign":"CQC_NBS","impressions":27,"clicks":10,"spend":15.1154,"conversions":0,"conversion_value":0},{"search_term":"club car seat covers","campaign":"CQC_NBS","impressions":39,"clicks":7,"spend":8.1721,"conversions":0,"conversion_value":0},{"search_term":"club car tempo seat covers","campaign":"CQC_NBS","impressions":11,"clicks":5,"spend":8.6537,"conversions":0,"conversion_value":0},{"search_term":"evolution golf cart seat covers","campaign":"CQC_NBS","impressions":12,"clicks":6,"spend":8.3048,"conversions":0,"conversion_value":0},{"search_term":"ezgo rxv seat covers","campaign":"CQC_NBS","impressions":11,"clicks":4,"spend":4.6892,"conversions":0,"conversion_value":0},{"search_term":"ezgo txt seat covers","campaign":"CQC_NBS","impressions":10,"clicks":4,"spend":4.7573,"conversions":0,"conversion_value":0},{"search_term":"denago seat covers","campaign":"CQC_NBS","impressions":5,"clicks":3,"spend":6.74,"conversions":0,"conversion_value":0},{"search_term":"denago rover xl seat cover","campaign":"CQC_NBS","impressions":3,"clicks":3,"spend":4.3312,"conversions":0,"conversion_value":0},{"search_term":"custom golf cart seat covers","campaign":"CQC_NBS","impressions":26,"clicks":6,"spend":4.6033,"conversions":0,"conversion_value":0},{"search_term":"club car onward seat covers","campaign":"CQC_NBS","impressions":13,"clicks":3,"spend":4.8295,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat covers","campaign":"CQC_NBS","impressions":33,"clicks":4,"spend":6.51,"conversions":0,"conversion_value":0},{"search_term":"club car precedent seat covers","campaign":"CQC_NBS","impressions":10,"clicks":3,"spend":3.4128,"conversions":0,"conversion_value":0},{"search_term":"ezgo golf cart seat covers","campaign":"CQC_NBS","impressions":9,"clicks":3,"spend":1.7522,"conversions":0,"conversion_value":0},{"search_term":"icon golf cart seat covers","campaign":"CQC_NBS","impressions":4,"clicks":3,"spend":2.2564,"conversions":0,"conversion_value":0},{"search_term":"club car carryall seat covers","campaign":"CQC_NBS","impressions":4,"clicks":2,"spend":3.5388,"conversions":0,"conversion_value":0},{"search_term":"club car golf cart seat covers","campaign":"CQC_NBS","impressions":9,"clicks":2,"spend":2.8495,"conversions":0,"conversion_value":0},{"search_term":"seat covers for club car precedent","campaign":"CQC_NBS","impressions":4,"clicks":2,"spend":1.4424,"conversions":0,"conversion_value":0},{"search_term":"ez go golf cart seat covers","campaign":"CQC_NBS","impressions":7,"clicks":1,"spend":4.55,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers ezgo rxv","campaign":"CQC_NBS","impressions":1,"clicks":2,"spend":2.18,"conversions":0,"conversion_value":0},{"search_term":"club car ds seat covers","campaign":"CQC_NBS","impressions":9,"clicks":3,"spend":3.25,"conversions":0,"conversion_value":0},{"search_term":"ezgo rxv golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat covers for golf carts","campaign":"CQC_NBS","impressions":10,"clicks":1,"spend":2.7523,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive 2 seat covers","campaign":"CQC_NBS","impressions":7,"clicks":1,"spend":0.7441,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive seat covers","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.65,"conversions":0,"conversion_value":0},{"search_term":"yamaha g29 seat covers","campaign":"CQC_NBS","impressions":10,"clicks":1,"spend":1.63,"conversions":0,"conversion_value":0},{"search_term":"club car seat cover replacement","campaign":"CQC_NBS","impressions":3,"clicks":1,"spend":1.2149,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat cover replacement","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":0.84,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat cover replacement","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":0.85,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers club car","campaign":"CQC_NBS","impressions":26,"clicks":3,"spend":3.9964,"conversions":0,"conversion_value":0},{"search_term":"seat covers for yamaha golf carts","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat cover","campaign":"CQC_NBS","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g22 seat covers","campaign":"CQC_NBS","impressions":6,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for e z go","campaign":"CQC_NBS","impressions":11,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat cover","campaign":"CQC_NBS","impressions":21,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"denago golf cart seat covers","campaign":"CQC_NBS","impressions":5,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car seat covers front and back","campaign":"CQC_NBS","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car precedent seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"custom golf cart seat covers","campaign":"CQC_NBS","impressions":26,"clicks":6,"spend":4.6033,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for ez go","campaign":"CQC_NBS","impressions":3,"clicks":2,"spend":4.44,"conversions":0,"conversion_value":0},{"search_term":"ezgo marathon seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"custom seat covers for golf carts","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":3.835,"conversions":0,"conversion_value":0},{"search_term":"club car precedent front and rear seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive 2 golf cart seat covers","campaign":"CQC_NBS","impressions":6,"clicks":1,"spend":1.46,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf seat covers","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":1.29,"conversions":0,"conversion_value":0},{"search_term":"yamaha seat covers for golf carts","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":1.24,"conversions":0,"conversion_value":0},{"search_term":"seat covers for a yamaha golf cart","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.53,"conversions":0,"conversion_value":0},{"search_term":"seat covers for yamaha golf cart","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.01,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers ezgo","campaign":"CQC_NBS","impressions":7,"clicks":2,"spend":3.13,"conversions":0,"conversion_value":0},{"search_term":"club car golf seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf car seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ezgo seat covers","campaign":"CQC_NBS","impressions":8,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"custom golf seat covers","campaign":"CQC_NBS","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ezgo txt seat cover replacement","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":1.4992,"conversions":0,"conversion_value":0},{"search_term":"golf seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat cover for golf cart","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":0.5392,"conversions":0,"conversion_value":0},{"search_term":"seat covers for evolution golf cart","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.42,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart replacement seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf seat covers club car","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.9018,"conversions":0,"conversion_value":0},{"search_term":"club car precedent seat covers front and back","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.17,"conversions":0,"conversion_value":0},{"search_term":"custom made golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart custom seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"custom golf cart seat cover","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":1.77,"conversions":0,"conversion_value":0},{"search_term":"golf cart rear seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"evolution golf cart seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf carts seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seats covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"universal golf cart seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seats and covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"covers for golf cart seats","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat cover for club car golf cart","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ezgo replacement seat covers","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":2.16,"conversions":0,"conversion_value":0},{"search_term":"yamaha g9 seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g29 golf cart seat covers","campaign":"CQC_NBS","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g22 seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g29 seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers staple on","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf seats covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart vinyl seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat upholstery","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat covers for golf cart","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"2 seat golf cart cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ez go seat covers","campaign":"CQC_NBS","impressions":6,"clicks":1,"spend":0.87,"conversions":0,"conversion_value":0},{"search_term":"e z go seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"rxv seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ezgo seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"denago golf cart covers","campaign":"CQC_NBS","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car ds seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club cart precedent seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers florida","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"personalized golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"collegiate golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car ds seat covers front and rear","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for club car","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat covers for club car golf cart","campaign":"CQC_NBS","impressions":4,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat covers for club car golf carts","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car golf cart seat cover","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":1.62,"conversions":0,"conversion_value":0},{"search_term":"1999 club car ds seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"pre 2000 club car ds seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart accessories seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for sale","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat replacement covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"ez golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"icon seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"etsy golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car golf cart seat covers","campaign":"CQC_NBS","impressions":9,"clicks":2,"spend":2.8495,"conversions":0,"conversion_value":0},{"search_term":"yamaha g16 seat cover","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g16 seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers near me","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car precedent golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g2 seat covers","campaign":"CQC_NBS","impressions":2,"clicks":1,"spend":0.38,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat blanket","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha seat cover replacement","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha g9 seat cover","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers club car ds","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"club car precedent front seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"evolution seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart slip on seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive golf cart seat covers","campaign":"CQC_NBS","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for ezgo","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"seat covers for ezgo txt golf cart","campaign":"CQC_NBS","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"star golf cart seat covers","campaign":"CQC_NBS","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"best golf cart seat covers","campaign":"CQC_NBS","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers yamaha","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":1.5441,"conversions":0,"conversion_value":0},{"search_term":"seat cover for golf cart","campaign":"CQC_NBS","impressions":1,"clicks":1,"spend":0.5392,"conversions":0,"conversion_value":0}],
"Ripco": [{"search_term":"behavioral furniture","campaign":"RipCo Products","impressions":1,"clicks":1,"spend":4.5424,"conversions":0,"conversion_value":0},{"search_term":"jail bed","campaign":"RipCo Products","impressions":4,"clicks":1,"spend":3.61,"conversions":1,"conversion_value":0},{"search_term":"prison beds","campaign":"RipCo Products","impressions":4,"clicks":1,"spend":3.96,"conversions":0,"conversion_value":0},{"search_term":"prison furniture","campaign":"RipCo Products","impressions":15,"clicks":2,"spend":4.94,"conversions":0,"conversion_value":0},{"search_term":"prison mattress","campaign":"RipCo Products","impressions":13,"clicks":1,"spend":2.75,"conversions":0,"conversion_value":0},{"search_term":"institutional mattress","campaign":"RipCo Products","impressions":1,"clicks":1,"spend":2.1,"conversions":0,"conversion_value":0},{"search_term":"correctional industries furniture","campaign":"RipCo Products","impressions":1,"clicks":1,"spend":2.42,"conversions":0,"conversion_value":0},{"search_term":"correctional mattresses","campaign":"RipCo Products","impressions":2,"clicks":1,"spend":3.82,"conversions":0,"conversion_value":0},{"search_term":"prison mattresses","campaign":"RipCo Products","impressions":1,"clicks":1,"spend":3.82,"conversions":0,"conversion_value":0},{"search_term":"behavioral health furniture","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"detention furniture","campaign":"RipCo Products","impressions":6,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"corrections furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail furniture","campaign":"RipCo Products","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"inmate furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail beds","campaign":"RipCo Products","impressions":9,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison bed","campaign":"RipCo Products","impressions":13,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail mattress","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison bunk","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison bunk beds","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"inmate bunk beds","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"mental health furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"furniture for behavioral health units","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"behavioral healthcare furniture","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"psych safe furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"beds in jail","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"inmate mattress","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail bed mat","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail bunk beds","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail bunks","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail cot","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail mats","campaign":"RipCo Products","impressions":3,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison made furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison mats","campaign":"RipCo Products","impressions":2,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"psych ward furniture","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"cot in jail","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"inmate mattresses","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail bed mattress","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"jail cell mattress","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0},{"search_term":"prison cell beds","campaign":"RipCo Products","impressions":1,"clicks":0,"spend":0,"conversions":0,"conversion_value":0}]
}

# ── disk file mapping ────────────────────────────────────────────────────────
DISK_FILES = {
    "LEARN":      "mcp-Windsor-ai-get_data-1778595146512.txt",
    "Griffin":    "mcp-Windsor-ai-get_data-1778595147876.txt",
    "GoGather":   "mcp-Windsor-ai-get_data-1778595151750.txt",
    "EcoPeaCo US":"mcp-Windsor-ai-get_data-1778595152128.txt",
    "RCM":        "mcp-Windsor-ai-get_data-1778595152469.txt",
    "EcoPeaCo CA":"mcp-Windsor-ai-get_data-1778595153127.txt",
    "We Buy":     "mcp-Windsor-ai-get_data-1778595156118.txt",
    "Spotlight":  "mcp-Windsor-ai-get_data-1778595166410.txt",
    "ASRV":       "mcp-Windsor-ai-get_data-1778595360601.txt",
}

# ── tab order ────────────────────────────────────────────────────────────────
TAB_ORDER = ["ASRV","Griffin","Velo","Sunrise","TriActive","LEARN","Eurolog","Spotlight",
             "CQC","EcoPeaCo CA","EcoPeaCo US","GoGather","RCM","Ripco","We Buy"]

CLIENT_DISPLAY = {
    "ASRV":"ASRV","Griffin":"Griffin Funding","Velo":"Velo Property Solutions",
    "Sunrise":"Sunrise Yacht Products","TriActive":"TriActiveUSA","LEARN":"LEARN Behavioral",
    "Eurolog":"Eurolog Packing Group","Spotlight":"Spotlight.Vegas",
    "CQC":"Custom Quality Covers","EcoPeaCo CA":"EcoPeaCo Canada",
    "EcoPeaCo US":"EcoPeaCo USA","GoGather":"GoGather",
    "RCM":"Renewable Credit Management","Ripco":"Ripco Products","We Buy":"We Buy Exotics"
}

# ── load data ────────────────────────────────────────────────────────────────
def load_raw(client_key):
    if client_key in NEW_DATA:
        return NEW_DATA[client_key]
    fname = DISK_FILES.get(client_key)
    if not fname:
        return []
    path = os.path.join(DATA_DIR, fname)
    with open(path) as f:
        data = json.load(f)
    return data.get("result", [])

# ── threshold filter + dedup ─────────────────────────────────────────────────
def filter_and_dedup(rows):
    agg = defaultdict(lambda: {"impressions":0,"clicks":0,"spend":0.0,"conversions":0.0,"conversion_value":0.0,"campaigns":set()})
    for r in rows:
        st = r.get("search_term")
        if not st:
            continue
        agg[st]["impressions"]      += r.get("impressions", 0)
        agg[st]["clicks"]           += r.get("clicks", 0)
        agg[st]["spend"]            += r.get("spend", 0.0)
        agg[st]["conversions"]      += r.get("conversions", 0.0)
        agg[st]["conversion_value"] += r.get("conversion_value", 0.0)
        camp = r.get("campaign","")
        if camp:
            agg[st]["campaigns"].add(camp)
    # apply threshold: keep if clicks>=1 OR spend>=1.0
    filtered = []
    for st, m in agg.items():
        if m["clicks"] >= 1 or m["spend"] >= 1.0:
            filtered.append({
                "search_term": st,
                "impressions": m["impressions"],
                "clicks": m["clicks"],
                "spend": round(m["spend"], 2),
                "conversions": m["conversions"],
                "conversion_value": round(m["conversion_value"], 2),
                "campaigns": ", ".join(sorted(m["campaigns"])),
            })
    return filtered

# ── categorization ───────────────────────────────────────────────────────────
def contains_any(term, patterns):
    t = term.lower()
    return any(p.lower() in t for p in patterns)

def starts_with_any(term, patterns):
    t = term.lower()
    return any(t.startswith(p.lower()) for p in patterns)

JOB_TERMS = ["job","jobs","career","careers","hiring","salary","how much do","how to become",
              "certification program","degree program","employment","resume","staffing agency"]
DIY_TERMS  = ["diy","how to make","how to build","homemade","self-made","pattern","tutorial",
              "plans for","build your own","make your own","step by step"]
INFO_TERMS = ["what is ","how does ","why is ","history of ","definition of","explained","guide to ","overview of"]
CHEAP_TERMS= ["free ","cheap ","cheapest","budget ","discount code","coupon","promo code","how much does","price list"]

ACTIVEWEAR_COMPETITORS = ["lululemon","gymshark","under armour","nike","adidas","puma","reebok",
                           "vuori","rhone","ten thousand","cuts clothing","represent","stone island",
                           "alphalete","gym king","dsquared","acne studios"]

MORTGAGE_COMPETITORS   = ["rocket mortgage","quicken loans","better.com","better mortgage","penfed",
                           "home xpress","loanify","ps funding","johnson bank","royal united",
                           "guidefunding","integrity first","np inc","harbour realty","alphatech",
                           "splitero","loanDepot","loan depot","newrez","pennymac","chase bank home",
                           "amerisave","upstart"]

LV_NONSHOW_ELIMINATE   = ["hotel","resort","casino","flight","airline","restaurant","nightclub",
                           "strip club","tour ","food","brunch","buffet","spa ","pool party","club ",
                           "real estate","rent ","apartment","airbnb","vegas living","move to vegas",
                           "las vegas jobs","vegas jobs","jobs in las vegas","how to move","cheap flight",
                           "cheap hotel","budget hotel"]

def categorize(row, client_key, brand_terms, competitors, target_cpa):
    st   = row["search_term"].lower()
    spend= row["spend"]
    conv = row["conversions"]
    high_spend_threshold = target_cpa * 3

    # ── brand guard: never eliminate brand terms ──────────────────────────
    is_brand = contains_any(st, brand_terms) if brand_terms else False

    # ── competitor check ─────────────────────────────────────────────────
    if competitors and contains_any(st, competitors):
        return "COMPETITOR", "Appears to reference a competitor brand — may be intentional conquesting"

    # ─────────────────────── per-client logic ────────────────────────────
    if client_key == "ASRV":
        if contains_any(st, ACTIVEWEAR_COMPETITORS):
            return "COMPETITOR", "Contains a competing activewear brand name"
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — no commercial intent for purchasing apparel"
        if contains_any(st, DIY_TERMS):
            return "ELIMINATE", "DIY/make-it-yourself query — searcher is not a buyer"
        if "sewing" in st or "pattern" in st or "fabric" in st:
            return "ELIMINATE", "Sewing/fabric-making query — searcher is making, not buying"
        if contains_any(st, CHEAP_TERMS) and not is_brand:
            return "CONSIDER", "Price-sensitive modifier may indicate budget buyer misaligned with premium brand"
        if contains_any(st, INFO_TERMS) and not is_brand:
            return "CONSIDER", "Informational query — intent to purchase unclear"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — exceeds {high_spend_threshold:.0f} threshold and appears off-target"
        return None, None

    if client_key == "Griffin":
        if contains_any(st, MORTGAGE_COMPETITORS):
            return "COMPETITOR", "Contains a competing mortgage lender name — may be intentional conquesting"
        if contains_any(st, JOB_TERMS) and not is_brand:
            return "ELIMINATE", "Job-seeking query — no loan application intent"
        if "commercial real estate loan" in st or "commercial mortgage" in st or "commercial property loan" in st:
            return "ELIMINATE", "Commercial real estate — Griffin does not offer commercial loans"
        if "eb5" in st:
            return "ELIMINATE", "EB-5 investor visa loan — outside Griffin's product set"
        if "sba loan" in st and "real estate" not in st and "multifamily" not in st:
            return "ELIMINATE", "SBA loan — not Griffin's product; may indicate government-backed small business loan seeker"
        if "hud 184" in st or "hud loan" in st:
            return "ELIMINATE", "HUD 184 Native American loan — outside Griffin's product set"
        if "flip loan" in st or "hard money" in st:
            return "CONSIDER", "Fix-and-flip / hard money loan — adjacent but not Griffin's core product set"
        if "free" in st and "credit" in st:
            return "ELIMINATE", "Free credit check query — not a loan applicant"
        if "free" in st and "report" in st:
            return "ELIMINATE", "Credit report query — informational, not a mortgage applicant"
        if contains_any(st, INFO_TERMS) and not is_brand:
            return "CONSIDER", "Informational query — educational intent rather than immediate loan application"
        if "calculate" in st or "calculator" in st:
            return "CONSIDER", "Calculator/research query — likely early-funnel research, not high-intent applicant"
        if "can i use heloc" in st and "invest in stocks" in st:
            return "ELIMINATE", "HELOC-for-stock-investment query — outside Griffin's use-case and risky intent signal"
        if "interest only heloc calculator" in st:
            return "CONSIDER", "Interest-only HELOC calculator — research intent; lower purchase readiness"
        if "bond loan" in st:
            return "ELIMINATE", "Bond loan — unrelated financial product, not a mortgage"
        if "tax lien" in st:
            return "CONSIDER", "Tax lien loan — adjacent financial product, unclear if aligned with Griffin's offerings"
        if "500 000 loan" in st or "borrow 150k" in st:
            return "CONSIDER", "Generic large-loan query without mortgage-specific intent"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — worth reviewing for match quality"
        return None, None

    if client_key == "LEARN":
        if contains_any(st, JOB_TERMS) and not is_brand:
            return "ELIMINATE", "Job-seeking query — not a parent seeking therapy services"
        if "bcba" in st and any(x in st for x in ["salary","job","certification","exam","pass","credential","training course","become"]):
            return "ELIMINATE", "ABA professional credential/career query — targets practitioners, not families seeking services"
        if "aba therapist" in st and any(x in st for x in ["job","salary","hiring","work"]):
            return "ELIMINATE", "ABA therapist job search — not a family seeking services"
        if "little star aba" in st:
            return "COMPETITOR", "Little Star ABA — a competing ABA therapy provider"
        if "behavior analyst" in st and any(x in st for x in ["salary","certification","exam","program","degree"]):
            return "ELIMINATE", "Behavior analyst career/credential query — targets clinicians, not families"
        if contains_any(st, INFO_TERMS) and not is_brand:
            return "CONSIDER", "Informational query — may be a parent early in their research journey; worth monitoring"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend and apparent off-target intent"
        return None, None

    if client_key == "Spotlight":
        if contains_any(st, ["vegas.com", "vegascom"]):
            return "COMPETITOR", "Contains the primary competitor Vegas.com brand name — may be intentional conquesting"
        if contains_any(st, LV_NONSHOW_ELIMINATE) and not any(x in st for x in ["show","ticket","concert","comedy","magic","cirque","jabbawockeez","criss angel","blue man"]):
            return "ELIMINATE", "Non-show Las Vegas query — searcher not looking for show tickets"
        if any(x in st for x in ["flight","airline","airfare","fly to","flights to"]):
            return "ELIMINATE", "Airline/flight query — outside Spotlight's ticket product"
        if any(x in st for x in ["hotel","resort","motel","stay in vegas","where to stay","vegas accommodations"]):
            return "ELIMINATE", "Hotel/accommodation query — outside Spotlight's show-ticket product"
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — no ticket purchase intent"
        if contains_any(st, INFO_TERMS) and "las vegas" not in st:
            return "CONSIDER", "Generic informational query not specific to Las Vegas shows"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — review for show-ticket intent"
        return None, None

    if client_key == "Eurolog":
        # competitor brands
        if contains_any(st, ["absorgel","absortech","seadry","container dri","clariant","celluliner","climaliner","silverskin","topcod"]):
            return "COMPETITOR", "Contains a competing cargo protection product brand — may be intentional conquesting"
        # B2C/residential/DIY signals (CLAUDE.md specific warning)
        if any(x in st for x in ["home","house","backyard","diy","residential","living","shed","off grid","tiny home","conex home","container home","shipping container home","personal"]):
            return "ELIMINATE", "Consumer/residential-use query — EPGNA serves commercial export/B2B clients only"
        if "insulation for conex" in st:
            return "ELIMINATE", "Conex insulation for personal/home use — B2C query outside EPGNA's B2B scope"
        # generic moisture/condensation queries that could be B2C
        if "shipping container moisture" in st or "how to keep" in st or "how to stop" in st:
            return "CONSIDER", "Ambiguous intent — could be B2B shipper or B2C container-home user; verify via match type"
        return None, None

    if client_key == "Sunrise":
        if "treehousesupplies" in st:
            return "COMPETITOR", "Treehousesupplies.com — a competing online retailer for netting and rope products"
        if "boat for sale" in st or "catamaran for sale" in st or "leopard catamarans for sale" in st or "skipper 38 for sale" in st:
            return "ELIMINATE", "Boat purchase query — searcher is buying a vessel, not replacement netting"
        if st == "trampolines for sale" or (st == "trampoline" and "catamaran" not in st):
            return "ELIMINATE", "Generic backyard trampoline query — not a marine trampoline net buyer"
        if "lamborghini" in st:
            return "ELIMINATE", "Luxury car brand in sailing context — garbled/off-target query"
        if "trampoline material" in st:
            return "CONSIDER", "Material research query — could be a buyer evaluating fabric type or a DIYer"
        if "cargo net" in st and "obstacle" in st:
            return "ELIMINATE", "Obstacle-course cargo net — recreational equipment, not marine netting"
        if "outdoor loft netting" in st:
            return "CONSIDER", "Loft/architectural netting — adjacent product category; could be a legitimate lead or residential buyer"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — monitor for purchase intent"
        return None, None

    if client_key == "TriActive":
        if contains_any(st, ["hawkins dynamics","bertec","h e hodge","xplus outdoor"]):
            return "COMPETITOR", "Contains a competing outdoor/commercial fitness equipment brand"
        if "triactive america" in st:
            return "CONSIDER", "Possible competitor brand (TriActive America) or brand name variation — verify if different entity"
        if contains_any(st, DIY_TERMS):
            return "ELIMINATE", "DIY gym-build query — not a B2B/B2G commercial equipment buyer"
        if "kids" in st and "gym" in st:
            return "CONSIDER", "Children's gym equipment — may be residential/playground buyer, not TriActive's B2G target"
        if "used" in st:
            return "CONSIDER", "Used equipment query — not a new commercial purchase; likely misaligned with TriActive's new-equipment sales"
        if "for home" in st or "backyard" in st or "personal" in st or "home gym" in st:
            return "ELIMINATE", "Residential home gym query — TriActive targets B2B/B2G commercial buyers only"
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job/career query — not a commercial equipment buyer"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — generic term worth refining"
        return None, None

    if client_key == "Velo":
        if contains_any(st, ["go city wide","pro serve cleaning","cam resources","tssmc","greenhouse eco","rozalado","metro cleaning","ccs facility","mrt cleaning"]):
            return "COMPETITOR", "Contains a competing commercial cleaning company name — may be intentional conquesting"
        if contains_any(st, ["house cleaning","maid service","residential","home cleaning","apartment","airbnb","housekeeping"]):
            return "ELIMINATE", "Residential cleaning query — Velo is strictly B2B commercial; no residential services"
        if "franchise" in st:
            return "ELIMINATE", "Franchise opportunity query — searcher wants to buy a franchise, not hire a cleaning service"
        if contains_any(st, ["floor scrubber","pressure washer","industrial machine","cleaning equipment","cleaning supply"]) and "service" not in st:
            return "ELIMINATE", "Equipment purchase query — searcher wants to buy cleaning equipment, not hire a service"
        if "restaurant cleaning" in st or "restaurant kitchen" in st:
            return "CONSIDER", "Restaurant cleaning — adjacent commercial niche; Velo targets office/medical/warehouse, not food service"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — appears off-target"
        return None, None

    if client_key == "CQC":
        # Very niche account — golf cart seat covers only; minimal flags expected
        if "etsy" in st:
            return "CONSIDER", "Etsy marketplace search — buyer looking for handmade/marketplace options rather than a specialized supplier"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — review product-page experience"
        return None, None

    if client_key in ("EcoPeaCo CA", "EcoPeaCo US"):
        if contains_any(st, ACTIVEWEAR_COMPETITORS + ["pampers","huggies","gerber","carter","gap kids","old navy","h&m kids"]):
            return "COMPETITOR", "Contains a competing baby/children's apparel brand name"
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — no purchase intent"
        if contains_any(st, DIY_TERMS):
            return "ELIMINATE", "DIY/crafting query — not a buyer of finished eco baby products"
        if "100 cotton baby" in st or "organic cotton baby" in st:
            return "CONSIDER", "Material-focused query — could match any cotton baby brand; may be too broad"
        if "wholesale" in st or "bulk" in st:
            return "CONSIDER", "Wholesale/bulk query — may indicate a retailer buyer not aligned with DTC model"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — appears off-target for this DTC brand"
        return None, None

    if client_key == "GoGather":
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — not an event planner looking to book services"
        if contains_any(st, DIY_TERMS):
            return "ELIMINATE", "DIY event planning query — self-organizing rather than seeking professional services"
        if "free" in st and "event" in st:
            return "CONSIDER", "Free event query — may indicate someone planning a no-budget event, not a qualified B2B lead"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — appears off-target"
        return None, None

    if client_key == "RCM":
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — not a business seeking tax credit services"
        if "179d" == st.strip():
            return "CONSIDER", "Single tax-code query (179D) — could be a CPA/engineer researching, not a direct RCM buyer"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — appears off-target"
        return None, None

    if client_key == "Ripco":
        if "correctional industries" in st:
            return "COMPETITOR", "Correctional Industries programs manufacture and sell furniture — a competing supply channel for institutional buyers"
        if "prison made" in st:
            return "ELIMINATE", "Query seeking prison-manufactured furniture — refers to inmate-labor products, not commercial supplier"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "CONSIDER", f"Zero conversions with ${spend:.2f} spend — monitor for commercial buyer intent"
        return None, None

    if client_key == "We Buy":
        if contains_any(st, JOB_TERMS):
            return "ELIMINATE", "Job-seeking query — not a seller looking to sell an exotic vehicle"
        if "how to buy" in st or "buy a" in st:
            return "ELIMINATE", "Vehicle purchase query — We Buy Exotics buys cars; this is the opposite intent"
        if conv == 0 and spend > high_spend_threshold and not is_brand:
            return "ELIMINATE", f"Zero conversions with ${spend:.2f} spend — appears off-target"
        return None, None

    return None, None

# ── brand-term list per client ────────────────────────────────────────────────
BRAND_TERMS = {
    "ASRV":       ["asrv","a.s.r.v.","all season recreation"],
    "Griffin":    ["griffin funding","griffin lending","griffin mortgage","griffin"],
    "Velo":       ["velo property","velo janitorial","velo"],
    "Sunrise":    ["sunrise yacht","sunrise shade","sailor's pillow"],
    "TriActive":  ["triactive","tri-active","triactive usa"],
    "LEARN":      ["learn behavioral","learn academy"],
    "Eurolog":    ["eurolog","epgna","hybag","hyblanket","hypad","temcore"],
    "Spotlight":  ["spotlight.vegas","spotlight vegas"],
    "CQC":        [],
    "EcoPeaCo CA":["ecopeaco","eco pea co"],
    "EcoPeaCo US":["ecopeaco","eco pea co"],
    "GoGather":   ["gogather","go gather"],
    "RCM":        ["renewable credit management","rcm"],
    "Ripco":      ["ripco"],
    "We Buy":     ["we buy exotics"],
}

COMPETITORS_MAP = {
    "ASRV":        ACTIVEWEAR_COMPETITORS,
    "Griffin":     MORTGAGE_COMPETITORS,
    "Spotlight":   ["vegas.com","vegascom"],
    "Eurolog":     ["absorgel","absortech","seadry","container dri","clariant","celluliner","climaliner","silverskin","topcod"],
    "Sunrise":     ["treehousesupplies","intermas"],
    "TriActive":   ["hawkins dynamics","bertec","h e hodge","xplus outdoor"],
    "Velo":        ["go city wide","pro serve cleaning","cam resources","tssmc","greenhouse eco","rozalado","metro cleaning","ccs facility","mrt cleaning"],
    "LEARN":       ["little star aba"],
    "Ripco":       ["correctional industries"],
}

TARGET_CPA = {
    "ASRV":50,"Griffin":50,"Velo":50,"Sunrise":150,"TriActive":150,
    "LEARN":200,"Eurolog":500,"Spotlight":50,"CQC":50,
    "EcoPeaCo CA":50,"EcoPeaCo US":50,"GoGather":50,"RCM":50,"Ripco":50,"We Buy":50
}

# ── Excel helpers ─────────────────────────────────────────────────────────────
RED_FILL    = PatternFill("solid", fgColor="C00000")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL = PatternFill("solid", fgColor="ED7D31")
HEADER_FILL = PatternFill("solid", fgColor="2E4057")
SUMM_FILL   = PatternFill("solid", fgColor="D9E1F2")

WHITE_BOLD  = Font(bold=True, color="FFFFFF")
BLACK_BOLD  = Font(bold=True, color="000000")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT  = Font(bold=True, size=13)
SUMM_FONT   = Font(bold=True)

def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

def write_client_sheet(ws, client_key, flagged_rows, total_analyzed):
    display = CLIENT_DISPLAY[client_key]
    cat_counts = {"ELIMINATE":0,"CONSIDER":0,"COMPETITOR":0}
    for r in flagged_rows:
        cat_counts[r["category"]] += 1

    # Row 1 — merged header
    ws.merge_cells("A1:J1")
    ws["A1"] = f"{display} — Negative Keyword Report | Last 7 Days: {DATE_RANGE}"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Rows 3-6 — summary block
    summary = [
        ("Total Search Terms Analyzed", total_analyzed),
        ("Terms to Eliminate",          cat_counts["ELIMINATE"]),
        ("Terms to Consider",           cat_counts["CONSIDER"]),
        ("Competitor Terms",            cat_counts["COMPETITOR"]),
    ]
    for i, (label, val) in enumerate(summary, start=3):
        ws.cell(i, 1, label).font = SUMM_FONT
        ws.cell(i, 2, val).font   = SUMM_FONT
        ws.cell(i, 1).fill = SUMM_FILL
        ws.cell(i, 2).fill = SUMM_FILL

    if not flagged_rows:
        ws.cell(8, 1, "No terms flagged this week.")
        return

    # Row 8 — column headers
    headers = ["Category","Search Term","Matched Campaigns","Impressions","Clicks","Spend","Conversions","Conv. Value","Reason"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(8, col, h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    ws.freeze_panes = "A9"

    # sort: ELIMINATE → CONSIDER → COMPETITOR, then by spend desc
    order = {"ELIMINATE":0,"CONSIDER":1,"COMPETITOR":2}
    flagged_rows.sort(key=lambda x: (order[x["category"]], -x["spend"]))

    # data rows
    for row_idx, r in enumerate(flagged_rows, start=9):
        cat = r["category"]
        ws.cell(row_idx, 1, cat)
        ws.cell(row_idx, 2, r["search_term"])
        ws.cell(row_idx, 3, r["campaigns"])
        ws.cell(row_idx, 4, r["impressions"])
        ws.cell(row_idx, 5, r["clicks"])
        spend_cell = ws.cell(row_idx, 6, r["spend"])
        spend_cell.number_format = '$#,##0.00'
        ws.cell(row_idx, 7, r["conversions"])
        cv_cell = ws.cell(row_idx, 8, r["conversion_value"])
        cv_cell.number_format = '$#,##0.00'
        ws.cell(row_idx, 9, r["reason"])

        # color-code category cell
        cat_cell = ws.cell(row_idx, 1)
        if cat == "ELIMINATE":
            cat_cell.fill = RED_FILL
            cat_cell.font = WHITE_BOLD
        elif cat == "CONSIDER":
            cat_cell.fill = YELLOW_FILL
            cat_cell.font = BLACK_BOLD
        elif cat == "COMPETITOR":
            cat_cell.fill = ORANGE_FILL
            cat_cell.font = WHITE_BOLD
        cat_cell.alignment = Alignment(horizontal="center")

        ws.row_dimensions[row_idx].height = 16

    set_col_widths(ws, [14, 36, 40, 12, 8, 10, 13, 13, 60])

# ── main ─────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    summary_stats = {}

    for client_key in TAB_ORDER:
        print(f"Processing {client_key}...", end=" ")
        raw = load_raw(client_key)
        rows = filter_and_dedup(raw)
        total_analyzed = len(rows)

        brand_terms   = BRAND_TERMS.get(client_key, [])
        competitors   = COMPETITORS_MAP.get(client_key, [])
        target_cpa    = TARGET_CPA.get(client_key, 50)

        flagged = []
        for r in rows:
            if r["conversions"] > 0:
                # never eliminate converting terms
                cat, reason = categorize(r, client_key, brand_terms, competitors, target_cpa)
                if cat == "ELIMINATE":
                    cat, reason = "CONSIDER", reason + " (has conversions — downgraded from ELIMINATE)"
            else:
                cat, reason = categorize(r, client_key, brand_terms, competitors, target_cpa)
            if cat:
                flagged.append({**r, "category": cat, "reason": reason})

        ws = wb.create_sheet(title=CLIENT_DISPLAY[client_key][:31])
        write_client_sheet(ws, client_key, flagged, total_analyzed)

        cat_counts = {"ELIMINATE":0,"CONSIDER":0,"COMPETITOR":0}
        for fr in flagged:
            cat_counts[fr["category"]] += 1
        summary_stats[client_key] = cat_counts
        print(f"{total_analyzed} terms → E:{cat_counts['ELIMINATE']} C:{cat_counts['CONSIDER']} P:{cat_counts['COMPETITOR']}")

    out_path = f"/home/user/routines-ppc/reports/negative-keyword-report-{REPORT_DATE}.xlsx"
    wb.save(out_path)
    print(f"\nSaved: {out_path}")

    # print summary for Slack
    print("\n=== SUMMARY FOR SLACK ===")
    for client_key in TAB_ORDER:
        s = summary_stats[client_key]
        print(f"{client_key}: E={s['ELIMINATE']} C={s['CONSIDER']} P={s['COMPETITOR']}")

    return out_path, summary_stats

if __name__ == "__main__":
    main()
