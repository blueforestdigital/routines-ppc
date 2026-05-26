#!/usr/bin/env python3
"""
Weekly Negative Keyword Analysis Report Builder
Date: 2026-05-26
"""

import json
import os
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────

def filter_and_dedup(rows):
    """Apply threshold filter and deduplicate by search_term."""
    filtered = [r for r in rows
                if r.get('search_term')
                and (r.get('clicks', 0) >= 1 or r.get('spend', 0) >= 1.0)]
    dedup = {}
    for r in filtered:
        st = r['search_term']
        if st not in dedup:
            dedup[st] = {'search_term': st, 'campaigns': set(),
                         'impressions': 0, 'clicks': 0, 'spend': 0.0,
                         'conversions': 0.0, 'conversion_value': 0.0}
        dedup[st]['campaigns'].add(r.get('campaign', ''))
        dedup[st]['impressions'] += r.get('impressions', 0)
        dedup[st]['clicks'] += r.get('clicks', 0)
        dedup[st]['spend'] += r.get('spend', 0)
        dedup[st]['conversions'] += r.get('conversions', 0)
        dedup[st]['conversion_value'] += r.get('conversion_value', 0)
    return [{'search_term': k,
             'campaigns': ', '.join(sorted(v['campaigns'])),
             'impressions': v['impressions'],
             'clicks': v['clicks'],
             'spend': round(v['spend'], 2),
             'conversions': round(v['conversions'], 4),
             'conversion_value': round(v['conversion_value'], 2)}
            for k, v in dedup.items()]


def load_disk_file(path):
    with open(path) as f:
        data = json.load(f)
    if isinstance(data, dict) and 'result' in data:
        return filter_and_dedup(data['result'])
    elif isinstance(data, list):
        return data  # already processed
    return []


# Load disk files
BASE = '/root/.claude/projects/-home-user-routines-ppc/ebf8847d-9a5b-4eb5-a738-c6d4bc0e2819/tool-results/'
learn_data = load_disk_file(BASE + 'mcp-Windsor-ai-get_data-1779790207532.txt')
griffin_data = load_disk_file(BASE + 'mcp-Windsor-ai-get_data-1779790209090.txt')

with open('/home/user/routines-ppc/windsor_processed_output.json') as f:
    disk_data = json.load(f)
ecopeaco_ca_data = disk_data['EcoPeaCo Canada']
ecopeaco_us_data = disk_data['EcoPeaCo USA']
gogather_data = disk_data['GoGather']
rcm_data = disk_data['RCM']


# ─────────────────────────────────────────────
# INLINE RAW DATA (Eurolog, Sunrise, TriActive, Velo, CQC, Ripco, WBE)
# ─────────────────────────────────────────────

EUROLOG_RAW = [{"search_term":"container liners","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":6.7,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"dry bulk container liners","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":7.71,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"container desiccant bags","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":2.5896,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"desiccant bags","campaign":"BFD_NB_AllCampaigns","clicks":2,"spend":11.7,"impressions":47,"conversions":0,"conversion_value":0},{"search_term":"desiccant bags for shipping containers","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":4.93,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"desiccant for shipping containers","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":5.66,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"industrial insulation blankets","campaign":"BFD_NB_AllCampaigns","clicks":2,"spend":33.28,"impressions":14,"conversions":0,"conversion_value":0},{"search_term":"refrigerated shipping container insulation","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":6.91,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"thermal box panels","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":12.11,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"thermal blanket for food transportation","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":5.0888,"impressions":2,"conversions":1,"conversion_value":0},{"search_term":"pallet covers insulated","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":12.53,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"renew liner","campaign":"BFD_NB_AllCampaigns","clicks":2,"spend":16.7153,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"freezer blanket for food transport","campaign":"BFD_NB_AllCampaigns","clicks":3,"spend":26.0379,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"insulated food blankets","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":8.26,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"sea bulk container liner","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":11.38,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"container insulation kits","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":7.46,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"conex insulation","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":4.63,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"container insulation","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":1.06,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"container desiccant bag","campaign":"BFD_NB_AllCampaigns","clicks":1,"spend":5.596,"impressions":2,"conversions":0,"conversion_value":0}]

SUNRISE_RAW = [{"search_term":"hobie getaway trampoline","campaign":"multihull nets sales","clicks":1,"spend":2.09,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"multihull nets","campaign":"multihull nets sales","clicks":1,"spend":1.38,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"multihullnets","campaign":"multihull nets sales","clicks":2,"spend":4.31,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"multihull nets","campaign":"SMM_NBS_Catamaran_T2","clicks":2,"spend":0.0108,"impressions":3,"conversions":1,"conversion_value":0.75},{"search_term":"loft net","campaign":"multihull nets sales","clicks":1,"spend":1.03,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"multihullnets com","campaign":"multihull nets sales","clicks":1,"spend":0.01,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"loft nets","campaign":"multihull nets sales","clicks":1,"spend":1.22,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"multihullnets","campaign":"multihull nets sales","clicks":2,"spend":1.4808,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"catamaran nets","campaign":"multihull nets sales","clicks":1,"spend":0.51,"impressions":2,"conversions":1,"conversion_value":0.75},{"search_term":"trampoline bateau","campaign":"SMM_NBS_Catamaran_T2","clicks":1,"spend":1.97,"impressions":4,"conversions":1,"conversion_value":0.75},{"search_term":"loft hammock net","campaign":"multihull nets sales","clicks":1,"spend":0.28,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"loftnetz","campaign":"multihull nets sales","clicks":1,"spend":0.6171,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"multihullnets","campaign":"multihull nets sales","clicks":1,"spend":0.01,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"multihullnets","campaign":"multihull nets sales","clicks":1,"spend":0.01,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"dock net","campaign":"multihull nets sales","clicks":1,"spend":1.2,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"hammock floors","campaign":"multihull nets sales","clicks":1,"spend":0.62,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"net mesh","campaign":"multihull nets sales","clicks":1,"spend":0.73,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"loft nets","campaign":"multihull nets sales","clicks":1,"spend":1.22,"impressions":1,"conversions":0,"conversion_value":0}]

TRIACTIVE_RAW = [{"search_term":"prison exercise equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":0.66,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"outdoor exercise equipment for parks","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":14.66,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"outdoor fitness equipment","campaign":"SMM_NBS_Primary_KWs","clicks":7,"spend":15.3365,"impressions":53,"conversions":0,"conversion_value":0},{"search_term":"outdoor fitness equipment for parks","campaign":"SMM_NBS_Primary_KWs","clicks":2,"spend":7.0174,"impressions":13,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","clicks":15,"spend":41.7923,"impressions":100,"conversions":0,"conversion_value":0},{"search_term":"park exercise equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":1.89,"impressions":18,"conversions":0,"conversion_value":0},{"search_term":"weatherproof outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","clicks":3,"spend":5.49,"impressions":20,"conversions":0,"conversion_value":0},{"search_term":"outdoor fitness stations","campaign":"SMM_NBS_Primary_KWs","clicks":2,"spend":18.36,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"outdoor fitness trail equipment","campaign":"SMM_NBS_Primary_KWs","clicks":2,"spend":6.28,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym","campaign":"SMM_NBS_Primary_KWs","clicks":5,"spend":8.6703,"impressions":40,"conversions":0,"conversion_value":0},{"search_term":"triactive","campaign":"SMM_BS_TriActiveUSA","clicks":2,"spend":2.88,"impressions":15,"conversions":0,"conversion_value":0},{"search_term":"triactive usa fitness equipment","campaign":"SMM_BS_TriActiveUSA","clicks":3,"spend":3.6,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"triactive usa","campaign":"SMM_BS_TriActiveUSA","clicks":5,"spend":7.21,"impressions":11,"conversions":0,"conversion_value":0},{"search_term":"commercial outdoor fitness equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":3.84,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"exercise bike for outdoor use","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":1.23,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"outdoor fitness area","campaign":"SMM_NBS_Primary_KWs","clicks":0,"spend":0,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"park exercise machines","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":2.19,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"public park outdoor gym equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":8.21,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"kids outdoor fitness equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":1.6492,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"outdoor park exercise equipment","campaign":"SMM_NBS_Primary_KWs","clicks":3,"spend":16.4104,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"outdoor exercise equipment for adults","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":5.78,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"outdoor gym park equipment","campaign":"SMM_NBS_Primary_KWs","clicks":1,"spend":2.05,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"vuly outdoor gym","campaign":"SMM_NBS_Primary_KWs","clicks":0,"spend":0,"impressions":2,"conversions":0,"conversion_value":0}]

VELO_RAW = [{"search_term":"commercial floor scrubbers","campaign":"BFD_Velo_Google_NB_Search_Pros","clicks":1,"spend":9.6467,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"cleaners phoenix","campaign":"BFD_Velo_Google_NB_Search_Pros","clicks":1,"spend":20.02,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"floor cleaning zamboni","campaign":"BFD_Velo_Google_NB_Search_Pros","clicks":1,"spend":5.58,"impressions":1,"conversions":0,"conversion_value":0}]

CQC_RAW = [{"search_term":"club car seat covers","campaign":"CQC_NBS","clicks":6,"spend":5.2384,"impressions":29,"conversions":0,"conversion_value":0},{"search_term":"custom golf cart seat covers","campaign":"CQC_NBS","clicks":7,"spend":12.283,"impressions":22,"conversions":0,"conversion_value":0},{"search_term":"evolution golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.45,"impressions":20,"conversions":0,"conversion_value":0},{"search_term":"golf cart replacement seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.78,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers","campaign":"CQC_NBS","clicks":14,"spend":13.3846,"impressions":236,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers club car","campaign":"CQC_NBS","clicks":2,"spend":5.66,"impressions":24,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers evolution","campaign":"CQC_NBS","clicks":1,"spend":0.7192,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"golf seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.74,"impressions":9,"conversions":0,"conversion_value":0},{"search_term":"icon golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.91,"impressions":10,"conversions":0,"conversion_value":0},{"search_term":"yamaha g22 golf cart seat covers","campaign":"CQC_NBS","clicks":4,"spend":3.26,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat cover","campaign":"CQC_NBS","clicks":2,"spend":2.69,"impressions":14,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat covers","campaign":"CQC_NBS","clicks":4,"spend":5.05,"impressions":29,"conversions":0,"conversion_value":0},{"search_term":"club car golf cart seat covers","campaign":"CQC_NBS","clicks":2,"spend":2.2,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"club car onward seat covers","campaign":"CQC_NBS","clicks":1,"spend":2.72,"impressions":12,"conversions":0,"conversion_value":0},{"search_term":"club car precedent seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.71,"impressions":11,"conversions":0,"conversion_value":0},{"search_term":"ezgo golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.56,"impressions":7,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers ez go","campaign":"CQC_NBS","clicks":1,"spend":1.38,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers ezgo","campaign":"CQC_NBS","clicks":3,"spend":1.8411,"impressions":15,"conversions":0,"conversion_value":0},{"search_term":"denago golf cart seat covers","campaign":"CQC_NBS","clicks":5,"spend":7.965,"impressions":12,"conversions":0,"conversion_value":0},{"search_term":"denago seat covers","campaign":"CQC_NBS","clicks":2,"spend":3.2521,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"club car seat covers front and back","campaign":"CQC_NBS","clicks":1,"spend":0.73,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"custom made golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.62,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat upholstery","campaign":"CQC_NBS","clicks":1,"spend":1.27,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"icon seat covers","campaign":"CQC_NBS","clicks":3,"spend":4.56,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers for yamaha","campaign":"CQC_NBS","clicks":2,"spend":4.34,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers yamaha","campaign":"CQC_NBS","clicks":2,"spend":8.38,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"yamaha g29 seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.75,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat covers front and back","campaign":"CQC_NBS","clicks":2,"spend":1.8175,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"club car ds seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.94,"impressions":10,"conversions":0,"conversion_value":0},{"search_term":"club car premium seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.33,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"club car seat covers","campaign":"CQC_NBS","clicks":2,"spend":2.74,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat covers club car","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":7,"conversions":0,"conversion_value":0},{"search_term":"ez go golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.7,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"ez go seat covers","campaign":"CQC_NBS","clicks":2,"spend":6.32,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"ezgo rxv seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.14,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"golf cart seat upholstery","campaign":"CQC_NBS","clicks":1,"spend":1.27,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"staple on golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.18,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"yamaha golf cart seat cover replacement","campaign":"CQC_NBS","clicks":1,"spend":1.9584,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"club car ds golf cart seat covers","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"club car premium seat covers","campaign":"CQC_NBS","clicks":1,"spend":1.5287,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"seat covers for club car","campaign":"CQC_NBS","clicks":2,"spend":2.168,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"club cart seat covers","campaign":"CQC_NBS","clicks":2,"spend":1.2016,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"double take golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.5288,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive 2 golf cart seat covers","campaign":"CQC_NBS","clicks":1,"spend":0.88,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"yamaha seat covers for golf carts","campaign":"CQC_NBS","clicks":1,"spend":2.18,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"golf seat cover","campaign":"CQC_NBS","clicks":1,"spend":1.08,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"yamaha drive 2 seat covers","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"yamaha g29 seat cover","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"yamaha g16 seat covers","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"seat covers for yamaha golf cart","campaign":"CQC_NBS","clicks":0,"spend":0,"impressions":2,"conversions":0,"conversion_value":0}]

RIPCO_RAW = [{"search_term":"behavioral health furniture","campaign":"RipCo Products","clicks":1,"spend":3.12,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"behavioral healthcare furniture","campaign":"RipCo Products","clicks":1,"spend":4.614,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"psychiatric unit furniture","campaign":"RipCo Products","clicks":1,"spend":4.61,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"prison bunks","campaign":"RipCo Products","clicks":1,"spend":5.93,"impressions":3,"conversions":0,"conversion_value":0},{"search_term":"institutional furniture","campaign":"RipCo Products","clicks":1,"spend":2.45,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"jail bunk beds","campaign":"RipCo Products","clicks":1,"spend":2.49,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"jail mattresses","campaign":"RipCo Products","clicks":1,"spend":2.48,"impressions":1,"conversions":0,"conversion_value":0}]

WBE_RAW = [{"search_term":"we buy exotics","campaign":"SMM_BS_We_Buy_Exotics","clicks":42,"spend":812.93,"impressions":168,"conversions":1,"conversion_value":1000},{"search_term":"we buy exotics 1040 calle recodo san clemente ca 92673","campaign":"SMM_BS_We_Buy_Exotics","clicks":4,"spend":58.78,"impressions":28,"conversions":0,"conversion_value":0},{"search_term":"we buy exotics reviews","campaign":"SMM_BS_We_Buy_Exotics","clicks":3,"spend":79.3774,"impressions":30,"conversions":0,"conversion_value":0},{"search_term":"webuyexotics com","campaign":"SMM_BS_We_Buy_Exotics","clicks":2,"spend":37.25,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"we buy exotic cars","campaign":"SMM_BS_We_Buy_Exotics","clicks":2,"spend":58.31,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"webuyexotics","campaign":"SMM_BS_We_Buy_Exotics","clicks":2,"spend":12.04,"impressions":13,"conversions":0,"conversion_value":0},{"search_term":"exotic car dealer","campaign":"SMM_NBS_Exotic_Cars","clicks":6,"spend":125.7138,"impressions":80,"conversions":0,"conversion_value":0},{"search_term":"exotic car dealers","campaign":"SMM_NBS_Exotic_Cars","clicks":5,"spend":183.3876,"impressions":46,"conversions":0,"conversion_value":0},{"search_term":"exotic car dealership","campaign":"SMM_NBS_Exotic_Cars","clicks":14,"spend":322.4388,"impressions":214,"conversions":0,"conversion_value":0},{"search_term":"exotic cars dealership","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":53.36,"impressions":15,"conversions":0,"conversion_value":0},{"search_term":"exotic dealership","campaign":"SMM_NBS_Exotic_Cars","clicks":6,"spend":107.3,"impressions":50,"conversions":0,"conversion_value":0},{"search_term":"sell exotic car","campaign":"SMM_NBS_Exotic_Cars","clicks":4,"spend":195.61,"impressions":23,"conversions":2,"conversion_value":2000},{"search_term":"sell my exotic car","campaign":"SMM_NBS_Exotic_Cars","clicks":3,"spend":196.16,"impressions":34,"conversions":1,"conversion_value":100},{"search_term":"sell your exotic car","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":74.75,"impressions":2,"conversions":1,"conversion_value":1000},{"search_term":"who buys sports cars","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":78.31,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"sell lamborghini","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.8288,"impressions":5,"conversions":1,"conversion_value":1000},{"search_term":"911 buyers","campaign":"SMM_NBS_Exotic_Cars","clicks":6,"spend":406.88,"impressions":57,"conversions":0,"conversion_value":0},{"search_term":"buy my porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":133.64,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"porsche consignment","campaign":"SMM_NBS_Exotic_Cars","clicks":3,"spend":144.38,"impressions":23,"conversions":0,"conversion_value":0},{"search_term":"sell my porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":7,"spend":540.07,"impressions":41,"conversions":2,"conversion_value":800},{"search_term":"sell my porsche 911","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":75.56,"impressions":12,"conversions":0,"conversion_value":0},{"search_term":"ilusso","campaign":"SMM_NBS_Competitors","clicks":30,"spend":401.4551,"impressions":370,"conversions":5,"conversion_value":3100},{"search_term":"illuso cars","campaign":"SMM_NBS_Competitors","clicks":1,"spend":4.48,"impressions":25,"conversions":0,"conversion_value":0},{"search_term":"illuso palm beach","campaign":"SMM_NBS_Competitors","clicks":1,"spend":8.01,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"illusso","campaign":"SMM_NBS_Competitors","clicks":1,"spend":5.61,"impressions":38,"conversions":0,"conversion_value":0},{"search_term":"ilusso cars","campaign":"SMM_NBS_Competitors","clicks":2,"spend":81.58,"impressions":52,"conversions":1,"conversion_value":1000},{"search_term":"ilusso costa mesa","campaign":"SMM_NBS_Competitors","clicks":1,"spend":2.54,"impressions":11,"conversions":0,"conversion_value":0},{"search_term":"ilusso exotic car dealership","campaign":"SMM_NBS_Competitors","clicks":1,"spend":4.9799,"impressions":18,"conversions":0,"conversion_value":0},{"search_term":"ilusso houston","campaign":"SMM_NBS_Competitors","clicks":1,"spend":10.26,"impressions":29,"conversions":0,"conversion_value":0},{"search_term":"ilusso palm beach","campaign":"SMM_NBS_Competitors","clicks":2,"spend":70.1584,"impressions":28,"conversions":0,"conversion_value":0},{"search_term":"exotic car buyers","campaign":"SMM_NBS_Exotic_Cars","clicks":3,"spend":228.6857,"impressions":8,"conversions":0.7615,"conversion_value":76.1537},{"search_term":"sell porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":3,"spend":222.97,"impressions":13,"conversions":0,"conversion_value":0},{"search_term":"luxury car buyers","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":152.44,"impressions":8,"conversions":1,"conversion_value":1000},{"search_term":"sell audi q5","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":71.06,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"j&s auto haus","campaign":"SMM_NBS_Competitors","clicks":2,"spend":21.66,"impressions":6,"conversions":0.4557,"conversion_value":455.669},{"search_term":"j&s autohaus","campaign":"SMM_NBS_Competitors","clicks":1,"spend":7.6699,"impressions":39,"conversions":0,"conversion_value":0},{"search_term":"js autohaus","campaign":"SMM_NBS_Competitors","clicks":3,"spend":43.73,"impressions":11,"conversions":0,"conversion_value":0},{"search_term":"j and s autohaus","campaign":"SMM_NBS_Competitors","clicks":1,"spend":13.79,"impressions":24,"conversions":0,"conversion_value":0},{"search_term":"jsautohausgroup","campaign":"SMM_NBS_Competitors","clicks":2,"spend":24.02,"impressions":5,"conversions":0,"conversion_value":0},{"search_term":"sell dodge viper","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":11.33,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell my ferrari","campaign":"SMM_NBS_Exotic_Cars","clicks":0,"spend":0,"impressions":9,"conversions":0,"conversion_value":0},{"search_term":"porsche buyers","campaign":"SMM_NBS_Exotic_Cars","clicks":5,"spend":280.8078,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"best place to sell my porsche 911","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":152.08,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"buyers porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":53.11,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell porsche 911 turbo","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":157.9,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"selling porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":10.14,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell my rolls royce","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":76.45,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"excotic car trader","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":47.4983,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"exotic cars trader","campaign":"SMM_NBS_Exotic_Cars","clicks":2,"spend":155.74,"impressions":10,"conversions":0,"conversion_value":0},{"search_term":"sell my exotic","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":78.35,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell luxury car","campaign":"SMM_NBS_Exotic_Cars","clicks":4,"spend":307.27,"impressions":10,"conversions":1,"conversion_value":100},{"search_term":"sell my luxury car","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.49,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"selling a luxury car","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":78.82,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"exotic car consignment","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":60.5819,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"aston martin sell","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":6.164,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell my audi r8","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.62,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"porsche buyer","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.0008,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"i want to sell my porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":76.09,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"best place to sell my porsche","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.07,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"porsche 911 sell","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":78.02,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"j&s autohaus group","campaign":"SMM_NBS_Competitors","clicks":1,"spend":9.83,"impressions":1,"conversions":0,"conversion_value":0},{"search_term":"bmw m5 sell","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":22.14,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"sell my bentley","campaign":"SMM_NBS_Exotic_Cars","clicks":0,"spend":0,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"exotic cars dealer","campaign":"SMM_NBS_Exotic_Cars","clicks":3,"spend":78.74,"impressions":6,"conversions":0,"conversion_value":0},{"search_term":"exoctic car dealership","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":26.47,"impressions":2,"conversions":0,"conversion_value":0},{"search_term":"dodge viper sell","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":78.94,"impressions":4,"conversions":0,"conversion_value":0},{"search_term":"how to sell a ferrari","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":12.44,"impressions":8,"conversions":0,"conversion_value":0},{"search_term":"sell my audi r8","campaign":"SMM_NBS_Exotic_Cars","clicks":1,"spend":77.62,"impressions":2,"conversions":0,"conversion_value":0}]

# Process inline raw data
eurolog_data = filter_and_dedup(EUROLOG_RAW)
sunrise_data = filter_and_dedup(SUNRISE_RAW)
triactive_data = filter_and_dedup(TRIACTIVE_RAW)
velo_data = filter_and_dedup(VELO_RAW)
cqc_data = filter_and_dedup(CQC_RAW)
ripco_data = filter_and_dedup(RIPCO_RAW)
wbe_data = filter_and_dedup(WBE_RAW)

print("Data loaded:")
print(f"  LEARN: {len(learn_data)} | Griffin: {len(griffin_data)}")
print(f"  Eurolog: {len(eurolog_data)} | Sunrise: {len(sunrise_data)} | TriActive: {len(triactive_data)}")
print(f"  Velo: {len(velo_data)} | CQC: {len(cqc_data)} | Ripco: {len(ripco_data)} | WBE: {len(wbe_data)}")
print(f"  EcoPeaCo CA: {len(ecopeaco_ca_data)} | EcoPeaCo US: {len(ecopeaco_us_data)}")
print(f"  GoGather: {len(gogather_data)} | RCM: {len(rcm_data)}")


# ─────────────────────────────────────────────
# CATEGORIZATION
# ─────────────────────────────────────────────

def contains_any(term, keywords):
    t = term.lower()
    return any(kw.lower() in t for kw in keywords)

def starts_with_any(term, keywords):
    t = term.lower()
    return any(t.startswith(kw.lower()) for kw in keywords)

def exact_match_any(term, keywords):
    t = term.lower().strip()
    return any(kw.lower().strip() == t for kw in keywords)


def categorize_eurolog(row):
    st = row['search_term']
    # Competitors
    competitors = ['renew liner', 'vinliner', 'unifrax durablanket', 'clariant container dri',
                   'absorbpoles', 'absorgel', 'cargo drypak', 'topcod desiccant',
                   'infinity core thermal liner']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor brand / product name — review only'
    # Eliminate: DIY/residential container insulation
    eliminate_patterns = ['how to insulate', 'insulating a', 'insulation for shipping container home',
                          'shipping container insulation ideas', 'container insulation kits',
                          'conex box insulation', 'connex insulation', 'sea container insulation',
                          'shipping container outside insulation', 'cargo container insulation',
                          'conex container insulation', 'best way to insulate']
    if contains_any(st, eliminate_patterns):
        if row['conversions'] > 0:
            return 'CONSIDER', 'Has conversion — cannot eliminate; DIY/residential insulation query warrants further review'
        return 'ELIMINATE', 'DIY/residential container insulation search — not B2B industrial packaging audience'
    if 'insulation for a shipping container' in st.lower():
        return 'ELIMINATE', 'DIY container insulation query — wrong audience for B2B export packaging'
    if 'container insulation' == st.lower():
        return 'CONSIDER', 'Generic container insulation term — could attract DIY/residential traffic; monitor'
    return 'CONSIDER', 'Relevant industrial packaging / export product term — monitor for conversion'


def categorize_sunrise(row):
    st = row['search_term']
    brands = ['multihullnets', 'multihull nets', 'sunrise']
    if contains_any(st, brands) and 'SMM_' not in row.get('campaigns', ''):
        return 'CONSIDER', 'Brand / near-brand term — cannot recommend as negative'
    if contains_any(st, ['multihullnets', 'multihull nets']):
        return 'CONSIDER', 'Brand or branded product term — retain'
    # Competitors: specific boat/trampoline brands
    competitors = ['hobie', 'prindle', 'leopard 46', 'outremer', 'farrier', 'redden net',
                   'loftnets', 'loftnet']
    # hobie IS a competitor boat brand
    if contains_any(st, ['hobie', 'prindle', 'leopard 46', 'outremer', 'farrier', 'redden net']):
        return 'COMPETITOR', 'Competitor boat / trampoline brand — review only'
    # Eliminate
    eliminate = ['boat dock builders', 'mega yacht supply', 'copart boats',
                 'sail kit for inflatable', 'tree web']
    if contains_any(st, eliminate):
        return 'ELIMINATE', 'Off-topic / wrong product category for marine netting business'
    # Adjacent/loft netting products — Sunrise sells architectural loft nets
    loft = ['loft net', 'loftnets', 'loftnet', 'loftnetz', 'hammock floor', 'hammock net',
            'net loft', 'cargo net treehouse', 'treehouse', 'dock net', 'tree net',
            'lounging net', 'indoor net hammock']
    if contains_any(st, loft):
        return 'CONSIDER', 'Adjacent architectural / loft netting term — Sunrise sells loft nets; could be relevant audience'
    return 'CONSIDER', 'Marine netting / catamaran adjacent term — monitor for conversion'


def categorize_triactive(row):
    st = row['search_term']
    # Brand
    if contains_any(st, ['triactive', 'triactiveusa']):
        return 'CONSIDER', 'Client brand term — cannot recommend as negative'
    # Competitors
    competitors = ['vuly', 'life fitness outdoor', 'fitblock', 'paradigm', 'playcore',
                   'kompan', 'greenfields', 'biowheels']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor brand — review only'
    # Eliminate: consumer/residential, wrong geo, DIY
    eliminate = ['garden exercise', 'outdoor patio gym', 'personal outdoor gym',
                 'build your own outdoor', 'outdoor gym uk', 'outdoor gym equipment uk',
                 'kids outdoor gym equipment', 'outdoor dip', 'backyard',
                 'exercise bike for outdoor use', 'kids outdoor fitness']
    if contains_any(st, eliminate):
        if row['conversions'] > 0:
            return 'CONSIDER', 'Has conversion — cannot eliminate; review for residential vs institutional intent'
        return 'ELIMINATE', 'Consumer/residential fitness term — TriActive targets B2B/B2G institutional buyers'
    # Prison terms ARE relevant for TriActive
    if contains_any(st, ['prison', 'jail', 'correctional', 'detention']):
        return 'CONSIDER', 'Prison/correctional segment is a valid TriActive B2G target market'
    return 'CONSIDER', 'Commercial outdoor fitness term — relevant to parks/institutional segment'


def categorize_learn(row):
    st = row['search_term']
    # Brand terms for LEARN
    brand = ['bci worcester', 'bci learn behavioral', 'learn behavioral', 'behavioral concepts bci',
             'bci autism']
    if contains_any(st, brand):
        return 'CONSIDER', 'LEARN Behavioral / BCI brand term — cannot recommend as negative'
    # Competitors — specific ABA provider names
    competitors = ['kaleidoscope aba', 'kaleidoscope worcester', 'cortica', 'triangle aba',
                   'autism learning partners', 'above and beyond therapy', 'virtus health',
                   'circle city aba', 'hopebridge', 'bright bridge aba',
                   'autism center of excellence', 'greater boston aba',
                   'aba centers of america', 'amego', 'proven behavior solutions',
                   'golden steps aba', 'community autism services', 'key autism',
                   'proud moments aba', 'light house autism', 'journeys behavior learning center',
                   'kids count therapy', 'life strategies', 'all about aba indiana',
                   'amaze aba', 'applied behavior institute', 'aba with abc',
                   'proven autism therapies', 'autism center of excellence',
                   'autism specialist']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor ABA provider name — review only; do not add as negative'
    # Eliminate: wrong service type
    eliminate = ['speech therapy', 'occupational therapy', 'physical therapy', 'tutoring',
                 'daycare', 'preschool']
    if contains_any(st, eliminate) and not contains_any(st, ['autism', 'aba', 'behavioral']):
        return 'ELIMINATE', 'Different therapy/service type — not ABA therapy audience'
    # Spanish language terms — geo-targeted campaigns may not serve Spanish speakers
    if contains_any(st, ['psicologia', 'como funciona', 'como funciona']):
        return 'CONSIDER', 'Spanish-language query — review if language targeting is configured for geo campaigns'
    return 'CONSIDER', 'Autism / ABA therapy funnel term — relevant to LEARN Behavioral service offering'


def categorize_griffin(row):
    st = row['search_term']
    # Brand
    if contains_any(st, ['griffin funding', 'griffin lending']):
        return 'CONSIDER', 'Griffin Funding brand term — cannot recommend as negative'
    # Competitors
    competitors = ['rocket mortgage', 'quicken loans', 'loanDepot', 'united wholesale',
                   'better mortgage', 'pennymac', 'newrez', 'crosscountry', 'newfi',
                   'angel oak', 'civic financial', 'deephaven', 'acra lending',
                   'athas capital', 'velocity mortgage', 'lima one', 'kiavi', 'visio']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor mortgage lender — review only'
    # Eliminate: products Griffin doesn't offer
    eliminate_products = ['commercial loan', 'commercial mortgage', 'commercial real estate loan',
                          'sba loan', 'sba lending', 'business loan', 'business mortgage',
                          'construction loan', 'hard money loan', 'equipment loan',
                          'student loan', 'personal loan', 'auto loan', 'car loan',
                          'agricultural loan', 'farm loan']
    if contains_any(st, eliminate_products):
        if row['conversions'] > 0:
            return 'CONSIDER', 'Has conversion — cannot eliminate despite product mismatch; flag for review'
        return 'ELIMINATE', 'Loan product Griffin does not offer — wrong audience for DSCR/Non-QM/HELOC lender'
    # HELOC terms — flag NY/AK/MO geo if possible (can't geo-filter via negatives easily)
    # Keep as CONSIDER
    return 'CONSIDER', 'Mortgage / investment lending term — relevant to DSCR, Non-QM, BSL, or HELOC product lines'


def categorize_velo(row):
    st = row['search_term']
    # Brand
    if exact_match_any(st, ['velo']) or contains_any(st, ['velocity janitorial']):
        if exact_match_any(st, ['velo']):
            return 'CONSIDER', 'Velo brand term — cannot recommend as negative'
        return 'COMPETITOR', 'Competitor janitorial service — review only'
    # Competitors
    competitors = ['jani king', 'citywide', 'city wide', 'compass group cleaning',
                   'pentagon cleaning', 'omex cleaning', 'national cleaning',
                   'ace cleaning', 'bell rock cleaning', 'bullseye services',
                   'ccs commercial', 'corporate cleaning group', 'iss janitorial',
                   'jcs cleaning', 'kellermeyer', 'angel janitorial', 'cactus sweeping',
                   'velocity janitorial']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor commercial cleaning company — review only'
    # Eliminate: cleaning equipment (not services), residential, wrong service types
    equipment = ['floor scrubber', 'floor cleaner', 'cleaning zamboni', 'industrial floor cleaner',
                 'janilink', 'pringle floor', 'tennant q12', 'power washer industrial',
                 'pressure wash', 'pressure washing', 'power washing', 'az pressure']
    if contains_any(st, equipment):
        return 'ELIMINATE', 'Cleaning equipment or pressure washing — not commercial janitorial services; wrong product category'
    residential = ['residential cleaning', 'house cleaning', 'home cleaning', 'maid service',
                   'apartment cleaning', 'move out cleaning', 'move in cleaning']
    if contains_any(st, residential):
        return 'ELIMINATE', 'Residential cleaning — Velo targets commercial/institutional clients only'
    construction = ['post construction', 'construction cleanup', 'construction cleaning']
    if contains_any(st, construction):
        return 'ELIMINATE', 'Post-construction cleaning — specialized service Velo may not offer; confirm before excluding'
    return 'CONSIDER', 'Commercial janitorial service term — relevant to Velo\'s Phoenix/Scottsdale B2B market'


def categorize_cqc(row):
    st = row['search_term']
    # Competitor seat cover brands
    competitors = ['easy life golf cart', 'uptopar', 'double take golf cart',
                   'a to z golf cart', 'modz seat', 'cool dry golf cart',
                   'classic accessories golf cart']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor golf cart seat cover brand — review only'
    # Off-product: enclosures, rain covers, upholstery
    if contains_any(st, ['enclosure', 'rain cover']) and 'seat' not in st.lower():
        return 'ELIMINATE', 'Golf cart enclosure / rain cover — different product from seat covers; wrong product intent'
    if contains_any(st, ['upholstery']):
        return 'CONSIDER', 'Golf cart upholstery — adjacent to seat covers; could attract reupholstery searchers vs cover buyers'
    return 'CONSIDER', 'Golf cart seat cover term — relevant to CQC product line'


def categorize_ecopeaco(row):
    st = row['search_term']
    brand = ['eco pea', 'ecopeaco', 'eco peaco']
    if contains_any(st, brand):
        return 'CONSIDER', 'EcoPeaCo brand / near-brand term — cannot recommend as negative'
    # Competitors in baby/eco diaper space
    competitors = ['honest company', 'bambo nature', 'seventh generation', 'dyper',
                   'parasol', 'andy pandy', 'huggies', 'pampers', 'luvs',
                   'babyganics', 'naturally natures', 'coterie']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor diaper / baby product brand — review only'
    # Eliminate: clearly off-product
    eliminate = ['cat food', 'dog food', 'pet food', 'laundry', 'cleaning product',
                 'adult diaper', 'incontinence']
    if contains_any(st, eliminate):
        return 'ELIMINATE', 'Off-category product — not eco baby/diaper product'
    return 'CONSIDER', 'Eco baby product / diaper adjacent term — relevant to EcoPeaCo product line'


def categorize_gogather(row):
    st = row['search_term']
    brand = ['gogather']
    if contains_any(st, brand):
        return 'CONSIDER', 'GoGather brand term — cannot recommend as negative'
    competitors = ['cvent', 'eventbrite', 'bizzabo', 'whova', 'stova',
                   'allseated', 'tripleseat', 'etouches', 'aventri']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor event management platform — review only'
    # Eliminate: non-corporate/B2C events
    eliminate = ['birthday party', 'wedding', 'backyard party', 'kids party', 'graduation']
    if contains_any(st, eliminate):
        return 'ELIMINATE', 'Consumer/personal event — GoGather targets B2B corporate event planning'
    return 'CONSIDER', 'Corporate event planning / conference management term — relevant to GoGather services'


def categorize_rcm(row):
    st = row['search_term']
    brand = ['renewable credit management', 'rcm energy']
    if contains_any(st, brand):
        return 'CONSIDER', 'RCM brand term — cannot recommend as negative'
    # Eliminate: unrelated energy/financial terms
    eliminate = ['credit card', 'credit repair', 'credit score', 'tax credit personal',
                 'energy drink', 'electricity bill', 'utility bill', 'solar panel install']
    if contains_any(st, eliminate) and not contains_any(st, ['srec', 'rec ', 'renewable energy credit', 'carbon']):
        return 'ELIMINATE', 'Unrelated credit or energy term — not renewable energy credit trading'
    return 'CONSIDER', 'Renewable energy credit / SREC adjacent term — relevant to RCM trading services'


def categorize_ripco(row):
    st = row['search_term']
    # No known brand conflicts
    # Competitors in correctional/behavioral health furniture
    competitors = ['norix', 'american seating', 'republic furniture', 'discount store fixtures',
                   'national industries for the blind', 'preferred industries']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor institutional furniture brand — review only'
    # Eliminate: consumer furniture
    eliminate = ['home furniture', 'living room', 'bedroom furniture', 'office chair',
                 'dorm furniture', 'bunk bed' + 'for home', 'kids bunk bed']
    # All Ripco terms are institutional-relevant
    if contains_any(st, ['prison bunk bed', 'jail bunk bed']):
        return 'CONSIDER', 'Correctional facility bunk bed term — relevant to Ripco\'s correctional product line'
    return 'CONSIDER', 'Correctional / behavioral health institutional furniture term — relevant to Ripco product line'


def categorize_wbe(row):
    st = row['search_term']
    brand = ['we buy exotics', 'webuyexotics', 'we buy exotic cars', 'we buy exotics llc']
    if contains_any(st, brand):
        return 'CONSIDER', 'We Buy Exotics brand term — cannot recommend as negative'
    # Competitors
    competitors = ['ilusso', 'illusso', 'illuso', 'j&s autohaus', 'j&s auto haus',
                   'js autohaus', 'j and s autohaus', 'jsautohausgroup', 'j s autohaus',
                   'jands auto', 'j&j autohaus', 'jns auto haus', 'exotic auto trader',
                   'exoctic auto trader', 'excotic car trader', 'exotic car trader',
                   'motto renn', 'dupont registry', 'collect cars', 'cars and bids']
    if contains_any(st, competitors):
        return 'COMPETITOR', 'Competitor exotic car dealer / platform — review only'
    # Eliminate: buying-intent terms (WBE wants sellers, not buyers)
    # Actually WBE runs both buying and selling so "dealership" terms are CONSIDER
    # Eliminate very off-category terms
    eliminate = ['used car', 'budget car', 'economy car', 'car rental', 'lease car',
                 'car insurance', 'car repair', 'auto repair', 'car wash',
                 'cheap car', 'affordable car']
    if contains_any(st, eliminate):
        return 'ELIMINATE', 'Non-exotic / budget vehicle term — misaligned with WBE\'s high-end exotic car buying service'
    # Escalades / standard luxury SUVs — flag but CONSIDER
    if contains_any(st, ['escalade', 'q5', 'q7']) and not contains_any(st, ['exotic', 'luxury car buyers']):
        return 'CONSIDER', 'Standard luxury SUV term — WBE primarily buys exotics; confirm if they purchase this model'
    return 'CONSIDER', 'Exotic car buying / selling intent term — relevant to We Buy Exotics service'


def categorize_row(client_key, row):
    dispatch = {
        'Eurolog': categorize_eurolog,
        'Sunrise Yacht': categorize_sunrise,
        'TriActiveUSA': categorize_triactive,
        'LEARN Behavioral': categorize_learn,
        'Griffin Funding': categorize_griffin,
        'Velo': categorize_velo,
        'CQC': categorize_cqc,
        'EcoPeaCo Canada': categorize_ecopeaco,
        'EcoPeaCo USA': categorize_ecopeaco,
        'GoGather': categorize_gogather,
        'RCM': categorize_rcm,
        'Ripco': categorize_ripco,
        'We Buy Exotics': categorize_wbe,
    }
    fn = dispatch.get(client_key)
    if fn:
        return fn(row)
    return 'CONSIDER', 'Review required'


# Apply categorization
def categorize_dataset(client_key, rows):
    result = []
    for row in rows:
        cat, reason = categorize_row(client_key, row)
        row['category'] = cat
        row['reason'] = reason
        result.append(row)
    # Sort: ELIMINATE first, then CONSIDER, then COMPETITOR; within each by spend desc
    cat_order = {'ELIMINATE': 0, 'CONSIDER': 1, 'COMPETITOR': 2}
    return sorted(result, key=lambda r: (cat_order.get(r['category'], 3), -r['spend']))


all_clients = {
    'ASRV': {'account_id': '160-720-1558', 'data': [], 'no_data': True,
             'reason': 'Windsor.ai session expired — data unavailable for this reporting period'},
    'Eurolog': {'account_id': '187-108-3890', 'data': categorize_dataset('Eurolog', eurolog_data)},
    'Sunrise Yacht': {'account_id': '173-465-4183', 'data': categorize_dataset('Sunrise Yacht', sunrise_data)},
    'TriActiveUSA': {'account_id': '529-624-0197', 'data': categorize_dataset('TriActiveUSA', triactive_data)},
    'LEARN Behavioral': {'account_id': '723-428-3512', 'data': categorize_dataset('LEARN Behavioral', learn_data)},
    'Griffin Funding': {'account_id': '491-529-1526', 'data': categorize_dataset('Griffin Funding', griffin_data)},
    'Velo': {'account_id': '135-061-8133', 'data': categorize_dataset('Velo', velo_data)},
    'CQC': {'account_id': '584-530-6391', 'data': categorize_dataset('CQC', cqc_data)},
    'EcoPeaCo Canada': {'account_id': '930-498-4147', 'data': categorize_dataset('EcoPeaCo Canada', ecopeaco_ca_data)},
    'EcoPeaCo USA': {'account_id': '870-082-1596', 'data': categorize_dataset('EcoPeaCo USA', ecopeaco_us_data)},
    'GoGather': {'account_id': '834-239-8667', 'data': categorize_dataset('GoGather', gogather_data)},
    'RCM': {'account_id': '836-942-5452', 'data': categorize_dataset('RCM', rcm_data)},
    'Ripco': {'account_id': '742-528-0344', 'data': categorize_dataset('Ripco', ripco_data)},
    'We Buy Exotics': {'account_id': '611-293-3925', 'data': categorize_dataset('We Buy Exotics', wbe_data)},
    'Spotlight.Vegas': {'account_id': '835-493-5333', 'data': [], 'no_data': True,
                        'reason': 'Windsor.ai session expired — data unavailable for this reporting period'},
}

# Print summary
for name, info in all_clients.items():
    rows = info['data']
    if info.get('no_data'):
        print(f"  {name}: NO DATA")
    else:
        e = sum(1 for r in rows if r['category'] == 'ELIMINATE')
        c = sum(1 for r in rows if r['category'] == 'CONSIDER')
        comp = sum(1 for r in rows if r['category'] == 'COMPETITOR')
        print(f"  {name}: {len(rows)} terms | E:{e} C:{c} COMP:{comp}")


# ─────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────

REPORT_DATE = '2026-05-26'
DATE_FROM = '2026-05-19'
DATE_TO = '2026-05-25'

# Colors
FILL_ELIMINATE = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')   # light red
FILL_CONSIDER  = PatternFill(start_color='FFFACD', end_color='FFFACD', fill_type='solid')   # light yellow
FILL_COMPETITOR= PatternFill(start_color='FFE5CC', end_color='FFE5CC', fill_type='solid')   # light orange
FILL_HEADER    = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')   # dark blue
FILL_SUMMARY   = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')   # light blue

FONT_HEADER = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
FONT_TITLE  = Font(name='Calibri', bold=True, size=13)
FONT_LABEL  = Font(name='Calibri', bold=True, size=10)
FONT_NORMAL = Font(name='Calibri', size=10)
FONT_NODATA = Font(name='Calibri', italic=True, size=11, color='888888')

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

CAT_FILL = {
    'ELIMINATE': FILL_ELIMINATE,
    'CONSIDER':  FILL_CONSIDER,
    'COMPETITOR': FILL_COMPETITOR,
}

HEADERS = ['Search Term', 'Campaign(s)', 'Impr', 'Clicks', 'Spend ($)', 'Conv', 'Conv Value ($)', 'Category', 'Reason']
COL_WIDTHS = [40, 35, 8, 8, 11, 8, 14, 12, 55]


def write_client_sheet(ws, name, account_id, rows, no_data=False, no_data_reason=''):
    """Write a full client worksheet."""
    ws.sheet_view.showGridLines = True

    # Row 1: Title
    ws.row_dimensions[1].height = 22
    ws['A1'] = f'{name} — Negative Keyword Report — Week of {REPORT_DATE}'
    ws['A1'].font = FONT_TITLE
    ws.merge_cells('A1:I1')

    # Row 2: Account info
    ws['A2'] = f'Account: {account_id}  |  7-Day Window: {DATE_FROM} to {DATE_TO}'
    ws['A2'].font = FONT_LABEL
    ws.merge_cells('A2:I2')

    if no_data:
        # Just show no-data message
        ws['A4'] = '⚠  No data returned — skipping'
        ws['A4'].font = FONT_NODATA
        ws['A5'] = no_data_reason
        ws['A5'].font = FONT_NODATA
        ws.merge_cells('A4:I4')
        ws.merge_cells('A5:I5')
        return

    # Compute summary stats
    n_total = len(rows)
    n_elim = sum(1 for r in rows if r['category'] == 'ELIMINATE')
    n_cons = sum(1 for r in rows if r['category'] == 'CONSIDER')
    n_comp = sum(1 for r in rows if r['category'] == 'COMPETITOR')
    top_elim = sorted([r for r in rows if r['category'] == 'ELIMINATE'],
                      key=lambda r: -r['spend'])[:5]
    total_spend = sum(r['spend'] for r in rows)

    # Row 3: Counts
    ws['A3'] = f'Terms reviewed: {n_total}  |  ELIMINATE: {n_elim}  |  CONSIDER: {n_cons}  |  COMPETITOR: {n_comp}  |  Total 7d Spend: ${total_spend:,.2f}'
    ws['A3'].font = FONT_LABEL
    ws['A3'].fill = FILL_SUMMARY
    ws.merge_cells('A3:I3')

    # Row 4: Top eliminate terms
    if top_elim:
        top_str = '  |  '.join(f'"{r["search_term"]}" (${r["spend"]:.2f})' for r in top_elim)
        ws['A4'] = f'Top ELIMINATE by spend: {top_str}'
    else:
        ws['A4'] = 'Top ELIMINATE by spend: (none)'
    ws['A4'].font = FONT_NORMAL
    ws['A4'].fill = FILL_ELIMINATE
    ws.merge_cells('A4:I4')

    # Row 5: blank
    ws.row_dimensions[5].height = 6

    # Row 6: Headers
    for col_idx, (hdr, width) in enumerate(zip(HEADERS, COL_WIDTHS), start=1):
        cell = ws.cell(row=6, column=col_idx, value=hdr)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[6].height = 18

    # Freeze header row
    ws.freeze_panes = 'A7'

    # Data rows
    for row_idx, row in enumerate(rows, start=7):
        fill = CAT_FILL.get(row['category'], FILL_CONSIDER)
        values = [
            row['search_term'],
            row['campaigns'],
            row['impressions'],
            row['clicks'],
            row['spend'],
            row['conversions'],
            row['conversion_value'],
            row['category'],
            row['reason'],
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.font = FONT_NORMAL
            # Formatting
            if col_idx == 1:  # Search term
                cell.alignment = Alignment(wrap_text=False)
            elif col_idx == 2:  # Campaigns
                cell.alignment = Alignment(wrap_text=False)
            elif col_idx in (5, 7):  # Spend, Conv Value
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_idx in (3, 4, 6):  # Impr, Clicks, Conv
                cell.alignment = Alignment(horizontal='right')
            elif col_idx == 8:  # Category
                cell.alignment = Alignment(horizontal='center')
                cell.font = Font(name='Calibri', bold=True, size=10)
                cell.fill = fill
            elif col_idx == 9:  # Reason
                cell.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[row_idx].height = 15

    # Auto-fit row heights for reason column
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['I'].width = 55


def build_excel(output_path):
    wb = openpyxl.Workbook()
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Tab short names (max 31 chars)
    tab_names = {
        'ASRV': 'ASRV',
        'Eurolog': 'Eurolog',
        'Sunrise Yacht': 'Sunrise Yacht',
        'TriActiveUSA': 'TriActiveUSA',
        'LEARN Behavioral': 'LEARN Behavioral',
        'Griffin Funding': 'Griffin Funding',
        'Velo': 'Velo',
        'CQC': 'CQC',
        'EcoPeaCo Canada': 'EcoPeaCo Canada',
        'EcoPeaCo USA': 'EcoPeaCo USA',
        'GoGather': 'GoGather',
        'RCM': 'RCM',
        'Ripco': 'Ripco',
        'We Buy Exotics': 'We Buy Exotics',
        'Spotlight.Vegas': 'Spotlight.Vegas',
    }

    for name, info in all_clients.items():
        tab_name = tab_names[name]
        ws = wb.create_sheet(title=tab_name)
        write_client_sheet(
            ws, name, info['account_id'],
            info['data'],
            no_data=info.get('no_data', False),
            no_data_reason=info.get('reason', '')
        )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"\n✓ Excel saved: {output_path}")
    return output_path


OUTPUT = '/home/user/routines-ppc/reports/negative-keyword-report-2026-05-26.xlsx'
build_excel(OUTPUT)
print("Done!")
