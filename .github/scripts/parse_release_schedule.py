#!/usr/bin/env python3
"""
Parse release schedule markdown and generate branch names and CD parameters.
"""

import re
import json
import sys
from datetime import datetime
from typing import Dict, List, Optional


def extract_version_number(version_str: str) -> Optional[str]:
    """Extract version number from version string like 'v6.4.0' or '18.0 P2'."""
    if not version_str:
        return None
    
    # Remove common prefixes and clean up
    version_str = version_str.strip()
    
    # Match patterns like 6.4.0, 6.5, 18.0, etc.
    match = re.search(r'(\d+\.\d+(?:\.\d+)?)', version_str)
    if match:
        return match.group(1)
    
    return None


def determine_branch_name(product: str, version: str, release_type: str) -> Optional[str]:
    """
    Generate branch name based on product, version, and release type.
    
    VS Code: release/6.5 (for all 6.5.x versions)
    Visual Studio: release/VS180I4 (for 18.0 Insider 4)
    """
    if not version or version == '-':
        return None
    
    if 'VS Code' in product or product.strip() == 'VS Code':
        # Extract major.minor version for VS Code
        version_num = extract_version_number(version)
        if version_num:
            # Get major.minor only (e.g., 6.5 from 6.5.0)
            parts = version_num.split('.')
            if len(parts) >= 2:
                major_minor = f"{parts[0]}.{parts[1]}"
                return f"release/{major_minor}"
    
    elif 'VS' in product or 'Visual Studio' in product:
        # For Visual Studio, extract version and type
        # Examples: "18.0 P2" -> "release/VS180P2", "18.0 I5" -> "release/VS180I5"
        version_match = re.search(r'(\d+)\.(\d+)', version)
        type_match = re.search(r'(P|I|Preview|Insider)\s*(\d+)', version + ' ' + release_type)
        
        if version_match:
            major = version_match.group(1)
            minor = version_match.group(2)
            
            if type_match:
                type_letter = type_match.group(1)[0].upper()  # P or I
                type_num = type_match.group(2)
                return f"release/VS{major}{minor}{type_letter}{type_num}"
            else:
                # For GA releases without P/I designation
                return f"release/VS{major}{minor}"
    
    return None


def determine_preid(product: str, version: str, release_type: str) -> str:
    """
    Determine preid parameter for CD pipeline.
    
    VS Code:
    - Odd minor versions (6.3, 6.5): preview
    - Even minor versions (6.0, 6.2, 6.4): rc
    
    Visual Studio: preview
    """
    if 'VS Code' in product:
        version_num = extract_version_number(version)
        if version_num:
            parts = version_num.split('.')
            if len(parts) >= 2:
                try:
                    minor = int(parts[1])
                    # Odd minor version = preview, even = rc
                    if 'stable' in release_type.lower():
                        return 'rc' if minor % 2 == 0 else 'preview'
                    else:
                        return 'preview'
                except ValueError:
                    pass
        return 'preview'
    
    elif 'VS' in product or 'Visual Studio' in product:
        return 'preview'
    
    return 'preview'


def generate_series(product: str, version: str, branch_name: Optional[str], cut_date: str) -> str:
    """
    Generate series parameter.
    
    VS Code: CY260106 format (from cut date)
    Visual Studio: Same as branch name (e.g., VS180I4)
    """
    if 'VS Code' in product:
        # Parse cut date and generate CY format
        try:
            # Try various date formats
            for fmt in ['%b %d', '%b %d, %Y', '%m/%d', '%Y-%m-%d', '%m/%d/%Y']:
                try:
                    date_obj = datetime.strptime(cut_date.strip(), fmt)
                    # If year is not in the string, use current year
                    if date_obj.year == 1900:
                        date_obj = date_obj.replace(year=datetime.now().year)
                    return f"CY{date_obj.strftime('%y%m%d')}"
                except ValueError:
                    continue
        except Exception as e:
            print(f"Warning: Could not parse date '{cut_date}': {e}", file=sys.stderr)
        
        # Fallback to current date
        return f"CY{datetime.now().strftime('%y%m%d')}"
    
    elif 'VS' in product or 'Visual Studio' in product:
        if branch_name:
            # Remove 'release/' prefix
            return branch_name.replace('release/', '')
        return ''
    
    return ''


def determine_pkgs(product: str) -> str:
    """
    Determine pkgs parameter.
    
    Visual Studio: "server"
    VS Code: "" (blank)
    """
    if 'VS' in product and 'Code' not in product:
        return 'server'
    return ''


def determine_vsrelease(product: str) -> str:
    """
    Determine vsrelease parameter.
    
    Visual Studio: "true"
    Others: "false"
    """
    if 'VS' in product and 'Code' not in product:
        return 'true'
    return 'false'


def parse_markdown_table(content: str) -> List[Dict[str, str]]:
    """Parse markdown tables from the release schedule."""
    releases = []
    
    # Find all table rows (excluding header rows)
    lines = content.split('\n')
    in_table = False
    headers = []
    
    for line in lines:
        line = line.strip()
        
        # Check if this is a table row
        if '|' in line and line.startswith('|'):
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            
            # Skip empty rows
            if not any(cells):
                continue
            
            # Check if this is a header row
            if any(h.lower() in cell.lower() for cell in cells for h in ['Product', 'Release Type', 'Version', 'Cut Bits']):
                headers = cells
                in_table = True
                continue
            
            # Check if this is a separator row
            if all(set(cell.replace('-', '').strip()) <= {':'} for cell in cells if cell):
                continue
            
            # This is a data row
            if in_table and headers:
                row_dict = {}
                for i, cell in enumerate(cells):
                    if i < len(headers):
                        # Clean up cell content
                        clean_cell = re.sub(r'<[^>]+>', '', cell)  # Remove HTML tags
                        clean_cell = re.sub(r'\*\*([^*]+)\*\*', r'\1', clean_cell)  # Remove bold
                        clean_cell = re.sub(r'<br[^>]*>', ' ', clean_cell)  # Replace br with space
                        clean_cell = clean_cell.strip()
                        row_dict[headers[i]] = clean_cell
                
                # Only add if we have meaningful data
                if row_dict.get('Products') or row_dict.get('Product'):
                    releases.append(row_dict)
        elif in_table and not line.startswith('|'):
            # End of table
            in_table = False
            headers = []
    
    return releases


def generate_release_config(markdown_content: str) -> List[Dict]:
    """Generate release configuration from markdown schedule."""
    releases = parse_markdown_table(markdown_content)
    configs = []
    
    for release in releases:
        # Extract fields (try both singular and plural)
        product = release.get('Products') or release.get('Product', '')
        release_type = release.get('Release Type', '')
        version = release.get('Version', '')
        cut_date = release.get('Cut Bits Date', '')
        release_date = release.get('Release Date', '')
        status = release.get('Status', '')
        
        # Skip if no product or version
        if not product or not version or version == '-':
            continue
        
        # Skip if status indicates skip or already released
        if 'skip' in status.lower() or 'released' in status.lower() or 'canceled' in status.lower():
            continue
        
        # Generate branch name
        branch_name = determine_branch_name(product, version, release_type)
        if not branch_name:
            continue
        
        # Generate CD parameters
        config = {
            'product': product,
            'release_type': release_type,
            'version': version,
            'branch': branch_name,
            'cut_date': cut_date,
            'release_date': release_date,
            'cd_params': {
                'preid': determine_preid(product, version, release_type),
                'series': generate_series(product, version, branch_name, cut_date),
                'pkgs': determine_pkgs(product),
                'vsrelease': determine_vsrelease(product),
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
        
        # Output as JSON
        print(json.dumps(configs, indent=2))
        
    except FileNotFoundError:
        print(f"Error: File '{schedule_file}' not found", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
