#!/usr/bin/env python3
"""Release schedule parser (schedule-driven).

This script parses the markdown schedule and generates CD parameters.
All parameters must be explicitly provided in the schedule table.

Required columns per release row:
- Branch
- preid
- series
- pkgs
- vsrelease
"""

import re
import json
import sys
from datetime import datetime
from typing import Dict, List, Optional

def _lower_key_map(row: Dict[str, str]) -> Dict[str, str]:
    return {str(k).strip().lower(): (v if v is not None else "") for k, v in row.items()}


def _get_value(row: Dict[str, str], key: str) -> str:
    return _lower_key_map(row).get(key.strip().lower(), "").strip()


def _get_required_value(row: Dict[str, str], key: str, context: str) -> str:
    value = _get_value(row, key)
    if not value:
        raise ValueError(f"Missing required column '{key}' for {context}")
    return value


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


def generate_release_config(markdown_content: str) -> List[Dict]:
    """Generate release configuration from markdown schedule (schedule-driven)."""
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

        # Schedule-driven: all parameters must be explicitly provided
        context = f"{product} {version}".strip()
        try:
            branch = _get_required_value(release, 'branch', context)
            preid = _get_required_value(release, 'preid', context)
            series = _get_required_value(release, 'series', context)
            # pkgs can be empty string, but column must exist to make intent explicit
            _ = _get_required_value(release, 'pkgs', context)
            pkgs = _get_value(release, 'pkgs')
            vsrelease = _get_required_value(release, 'vsrelease', context)
        except ValueError as e:
            print(f"✗ Invalid schedule row: {e}", file=sys.stderr)
            continue
        
        config = {
            'product': product,
            'release_type': release_type,
            'version': version,
            'branch': branch,
            'cut_date': cut_date,
            'release_date': release_date,
            'cd_params': {
                'preid': preid,
                'series': series,
                'pkgs': pkgs,
                'vsrelease': vsrelease,
            }
        }
        
        configs.append(config)
    
    return configs


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Usage: parse_release_schedule.py <schedule.md>", file=sys.stderr)
        sys.exit(1)
    
    schedule_file = sys.argv[1]
    
    try:
        with open(schedule_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        configs = generate_release_config(content)
        
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
