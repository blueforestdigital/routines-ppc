#!/usr/bin/env python3
import json
import sys
from collections import defaultdict

files = [
    {
        "label": "EcoPeaCo Canada",
        "path": "/root/.claude/projects/-home-user-routines-ppc/ebf8847d-9a5b-4eb5-a738-c6d4bc0e2819/tool-results/mcp-Windsor-ai-get_data-1779790215166.txt"
    },
    {
        "label": "EcoPeaCo USA",
        "path": "/root/.claude/projects/-home-user-routines-ppc/ebf8847d-9a5b-4eb5-a738-c6d4bc0e2819/tool-results/mcp-Windsor-ai-get_data-1779790214628.txt"
    },
    {
        "label": "GoGather",
        "path": "/root/.claude/projects/-home-user-routines-ppc/ebf8847d-9a5b-4eb5-a738-c6d4bc0e2819/tool-results/mcp-Windsor-ai-get_data-1779790212925.txt"
    },
    {
        "label": "RCM",
        "path": "/root/.claude/projects/-home-user-routines-ppc/ebf8847d-9a5b-4eb5-a738-c6d4bc0e2819/tool-results/toolu_013NdcC3JGbN26bjxY9KJPj5.txt"
    },
]

def find_field(row, *candidates):
    """Try multiple field name candidates, return first match or None."""
    for c in candidates:
        if c in row:
            return row[c]
    return None

def safe_float(v):
    if v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0

def safe_int(v):
    if v is None:
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0

def process_file(label, path):
    with open(path, 'r', encoding='utf-8') as f:
        raw = f.read()

    # Try to parse as JSON directly
    data = json.loads(raw)

    # Data might be nested: {"result": [...]} or {"data": [...]} or just [...]
    if isinstance(data, dict):
        rows = data.get('result', data.get('data', data.get('rows', data.get('results', []))))
    elif isinstance(data, list):
        rows = data
    else:
        rows = []

    # Print sample row to understand field names
    if rows:
        print(f"# {label} - sample row keys: {list(rows[0].keys())}", file=sys.stderr)

    # Aggregate by search_term
    agg = defaultdict(lambda: {
        'impressions': 0,
        'clicks': 0,
        'spend': 0.0,
        'conversions': 0.0,
        'conversion_value': 0.0,
        'campaigns': set()
    })

    for row in rows:
        # Find search_term field
        search_term = find_field(row,
            'search_term', 'search term', 'query', 'search_query',
            'keyword', 'keyword_text', 'term'
        )
        if search_term is None:
            search_term = '__unknown__'

        impressions = safe_int(find_field(row, 'impressions', 'impression'))
        clicks = safe_int(find_field(row, 'clicks', 'click'))
        spend = safe_float(find_field(row, 'spend', 'cost', 'cost_micros'))
        conversions = safe_float(find_field(row, 'conversions', 'conversion', 'conversions_value'))
        conversion_value = safe_float(find_field(row,
            'conversion_value', 'conversions_value', 'revenue',
            'total_conversion_value', 'value'
        ))

        campaign = find_field(row,
            'campaign', 'campaign_name', 'campaign name',
            'ad_group', 'campaign_id'
        )

        # Filter: keep if clicks >= 1 OR spend >= 1.0
        if clicks >= 1 or spend >= 1.0:
            agg[search_term]['impressions'] += impressions
            agg[search_term]['clicks'] += clicks
            agg[search_term]['spend'] += spend
            agg[search_term]['conversions'] += conversions
            agg[search_term]['conversion_value'] += conversion_value
            if campaign:
                agg[search_term]['campaigns'].add(str(campaign))

    result = []
    for term, metrics in agg.items():
        result.append({
            'search_term': term,
            'campaigns': ', '.join(sorted(metrics['campaigns'])),
            'impressions': metrics['impressions'],
            'clicks': metrics['clicks'],
            'spend': round(metrics['spend'], 2),
            'conversions': round(metrics['conversions'], 2),
            'conversion_value': round(metrics['conversion_value'], 2),
        })

    return result

results = {}
for f in files:
    print(f"Processing {f['label']}...", file=sys.stderr)
    try:
        result = process_file(f['label'], f['path'])
        results[f['label']] = result
        print(f"  -> {len(result)} rows after filter+dedup", file=sys.stderr)
    except Exception as e:
        print(f"  ERROR: {e}", file=sys.stderr)
        import traceback; traceback.print_exc(file=sys.stderr)
        results[f['label']] = []

print(json.dumps(results, indent=2))
