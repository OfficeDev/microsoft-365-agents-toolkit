#!/usr/bin/env python3
"""
AI-powered release schedule parser using GitHub Models API.
This script uses LLM to intelligently parse release schedules and generate parameters.
"""

import re
import json
import sys
import os
from datetime import datetime
from typing import Dict, List, Optional
import urllib.request
import urllib.error


def call_github_model(prompt: str, model: str = "gpt-4o") -> str:
    """
    Call GitHub Models API to get intelligent parsing assistance.
    
    Args:
        prompt: The prompt to send to the model
        model: The model to use (default: gpt-4o)
    
    Returns:
        The model's response
    """
    github_token = os.environ.get('GITHUB_TOKEN')
    if not github_token:
        print("Warning: GITHUB_TOKEN not found, falling back to rule-based parsing", file=sys.stderr)
        return ""
    
    endpoint = "https://models.inference.ai.azure.com/chat/completions"
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {github_token}"
    }
    
    data = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "You are a release automation expert. Parse release schedules and generate precise configuration parameters in JSON format."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.1,
        "max_tokens": 2000
    }
    
    try:
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(data).encode('utf-8'),
            headers=headers
        )
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode('utf-8'))
            return result['choices'][0]['message']['content']
    
    except urllib.error.HTTPError as e:
        print(f"Warning: GitHub Models API error: {e.code} {e.reason}", file=sys.stderr)
        return ""
    except Exception as e:
        print(f"Warning: Failed to call GitHub Models API: {e}", file=sys.stderr)
        return ""


def extract_json_from_response(text: str) -> Optional[Dict]:
    """Extract JSON from LLM response that may contain markdown or other text."""
    # Try to find JSON in markdown code blocks
    json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except json.JSONDecodeError:
            pass
    
    # Try to find raw JSON
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(0))
        except json.JSONDecodeError:
            pass
    
    return None


def parse_with_ai(release_text: str, product: str, version: str, release_type: str, cut_date: str) -> Optional[Dict]:
    """
    Use AI to parse release information and generate parameters.
    
    Returns a dict with: branch, preid, series, pkgs, vsrelease
    """
    prompt = f"""
Parse the following release information and generate the configuration parameters:

**Product:** {product}
**Version:** {version}
**Release Type:** {release_type}
**Cut Date:** {cut_date}

Generate the following parameters following these rules:

1. **branch**: Branch name
   - VS Code: "release/X.Y" where X.Y is major.minor (e.g., "release/6.5" for version 6.5.0 or 6.5.20251203)
   - Visual Studio: "release/VSXXYTZ" where XX is major, Y is minor, T is type (P=Preview, I=Insider), Z is number
     Examples: "release/VS180P2" for "18.0 P2", "release/VS180I4" for "18.0 Insider 4"

2. **preid**: Release identifier
   - VS Code: "preview" if minor version is odd (6.3, 6.5) OR if release type contains "prerelease"
   - VS Code: "rc" if minor version is even (6.0, 6.2, 6.4, 6.6) AND release type is "Stable"
   - Visual Studio: always "preview"

3. **series**: Series name
   - VS Code: "CYYMMDD" format from cut date (e.g., "CY260106" for 2026-01-06)
   - Visual Studio: Same as branch name without "release/" prefix (e.g., "VS180I4")

4. **pkgs**: Package names
   - Visual Studio: "server"
   - VS Code: "" (empty string)

5. **vsrelease**: Boolean string
   - Visual Studio: "true"
   - VS Code: "false"

Return ONLY a JSON object with these fields (no additional text):
{{
  "branch": "...",
  "preid": "...",
  "series": "...",
  "pkgs": "...",
  "vsrelease": "..."
}}
"""
    
    response = call_github_model(prompt)
    if response:
        config = extract_json_from_response(response)
        if config and all(k in config for k in ['branch', 'preid', 'series', 'pkgs', 'vsrelease']):
            return config
    
    return None


def fallback_parse(product: str, version: str, release_type: str, cut_date: str) -> Optional[Dict]:
    """
    Fallback rule-based parsing if AI is not available.
    This preserves the original logic.
    """
    # Determine branch name
    branch = None
    if 'VS Code' in product or product.strip() == 'VS Code':
        version_match = re.search(r'(\d+)\.(\d+)', version)
        if version_match:
            major, minor = version_match.groups()
            branch = f"release/{major}.{minor}"
    elif 'VS' in product or 'Visual Studio' in product:
        version_match = re.search(r'(\d+)\.(\d+)', version)
        type_match = re.search(r'(P|I|Preview|Insider)\s*(\d+)', version + ' ' + release_type)
        if version_match:
            major, minor = version_match.groups()
            if type_match:
                type_letter = type_match.group(1)[0].upper()
                type_num = type_match.group(2)
                branch = f"release/VS{major}{minor}{type_letter}{type_num}"
            else:
                branch = f"release/VS{major}{minor}"
    
    if not branch:
        return None
    
    # Determine preid
    preid = "preview"
    if 'VS Code' in product:
        version_match = re.search(r'(\d+)\.(\d+)', version)
        if version_match:
            minor = int(version_match.group(2))
            if 'stable' in release_type.lower() and minor % 2 == 0:
                preid = "rc"
    
    # Determine series
    series = ""
    if 'VS Code' in product:
        try:
            for fmt in ['%b %d', '%b %d, %Y', '%m/%d', '%Y-%m-%d', '%m/%d/%Y']:
                try:
                    date_obj = datetime.strptime(cut_date.strip(), fmt)
                    if date_obj.year == 1900:
                        date_obj = date_obj.replace(year=datetime.now().year)
                    series = f"CY{date_obj.strftime('%y%m%d')}"
                    break
                except ValueError:
                    continue
        except:
            series = f"CY{datetime.now().strftime('%y%m%d')}"
    else:
        series = branch.replace('release/', '')
    
    # Determine pkgs and vsrelease
    pkgs = "server" if ('VS' in product and 'Code' not in product) else ""
    vsrelease = "true" if ('VS' in product and 'Code' not in product) else "false"
    
    return {
        'branch': branch,
        'preid': preid,
        'series': series,
        'pkgs': pkgs,
        'vsrelease': vsrelease
    }


def parse_markdown_table(content: str) -> List[Dict[str, str]]:
    """Parse markdown tables from the release schedule."""
    releases = []
    lines = content.split('\n')
    in_table = False
    headers = []
    
    for line in lines:
        line = line.strip()
        
        if '|' in line and line.startswith('|'):
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            
            if not any(cells):
                continue
            
            if any(h.lower() in cell.lower() for cell in cells for h in ['Product', 'Release Type', 'Version', 'Cut Bits']):
                headers = cells
                in_table = True
                continue
            
            if all(set(cell.replace('-', '').strip()) <= {':'} for cell in cells if cell):
                continue
            
            if in_table and headers:
                row_dict = {}
                for i, cell in enumerate(cells):
                    if i < len(headers):
                        clean_cell = re.sub(r'<[^>]+>', '', cell)
                        clean_cell = re.sub(r'\*\*([^*]+)\*\*', r'\1', clean_cell)
                        clean_cell = re.sub(r'<br[^>]*>', ' ', clean_cell)
                        clean_cell = clean_cell.strip()
                        row_dict[headers[i]] = clean_cell
                
                if row_dict.get('Products') or row_dict.get('Product'):
                    releases.append(row_dict)
        elif in_table and not line.startswith('|'):
            in_table = False
            headers = []
    
    return releases


def generate_release_config(markdown_content: str, use_ai: bool = True) -> List[Dict]:
    """Generate release configuration from markdown schedule."""
    releases = parse_markdown_table(markdown_content)
    configs = []
    
    for release in releases:
        product = release.get('Products') or release.get('Product', '')
        release_type = release.get('Release Type', '')
        version = release.get('Version', '')
        cut_date = release.get('Cut Bits Date', '')
        release_date = release.get('Release Date', '')
        status = release.get('Status', '')
        
        if not product or not version or version == '-':
            continue
        
        if 'skip' in status.lower() or 'released' in status.lower() or 'canceled' in status.lower():
            continue
        
        # Try AI-powered parsing first
        params = None
        if use_ai:
            params = parse_with_ai(markdown_content, product, version, release_type, cut_date)
            if params:
                print(f"✓ AI parsed: {product} {version}", file=sys.stderr)
        
        # Fallback to rule-based parsing
        if not params:
            params = fallback_parse(product, version, release_type, cut_date)
            if params:
                print(f"✓ Rule-based parsed: {product} {version}", file=sys.stderr)
        
        if not params or not params.get('branch'):
            print(f"✗ Could not parse: {product} {version}", file=sys.stderr)
            continue
        
        config = {
            'product': product,
            'release_type': release_type,
            'version': version,
            'branch': params['branch'],
            'cut_date': cut_date,
            'release_date': release_date,
            'cd_params': {
                'preid': params['preid'],
                'series': params['series'],
                'pkgs': params['pkgs'],
                'vsrelease': params['vsrelease'],
            }
        }
        
        configs.append(config)
    
    return configs


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Usage: parse_release_schedule.py <schedule.md> [--no-ai]", file=sys.stderr)
        sys.exit(1)
    
    schedule_file = sys.argv[1]
    use_ai = '--no-ai' not in sys.argv
    
    if use_ai and not os.environ.get('GITHUB_TOKEN'):
        print("Warning: GITHUB_TOKEN not set, using rule-based parsing only", file=sys.stderr)
        use_ai = False
    
    try:
        with open(schedule_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        configs = generate_release_config(content, use_ai=use_ai)
        
        print(json.dumps(configs, indent=2))
        
    except FileNotFoundError:
        print(f"Error: File '{schedule_file}' not found", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
