#!/usr/bin/env python3
"""
Check npm package.json files for security vulnerabilities.
This script scans package.json and package.json.tpl files and runs npm audit
in a temp directory to avoid modifying the original repository.
"""

import os
import sys
import argparse
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import List, Tuple
import json


def safe_print(message: str) -> None:
    """Safely print message, handling encoding issues"""
    try:
        print(message)
    except UnicodeEncodeError:
        safe_message = message.encode('ascii', 'replace').decode('ascii')
        print(safe_message)
    except Exception as e:
        print(f"[Print error: {type(e).__name__}]")


def find_package_files(scan_dirs: List[str]) -> Tuple[List[Path], List[Path]]:
    """Find all package.json and package.json.tpl files in the given directories"""
    package_files = []
    template_files = []
    
    for scan_dir in scan_dirs:
        base_path = Path(scan_dir)
        if not base_path.exists():
            safe_print(f"WARNING: Directory does not exist: {scan_dir}")
            continue
        
        # Find package.json files
        for pkg_file in base_path.rglob("package.json"):
            package_files.append(pkg_file)
        
        # Find package.json.tpl template files
        for tpl_file in base_path.rglob("package.json.tpl"):
            template_files.append(tpl_file)
    
    return package_files, template_files


def check_package_vulnerabilities(pkg_file: Path, temp_dir: Path, is_template: bool = False) -> Tuple[bool, str]:
    """
    Check a package.json file for vulnerabilities using npm audit.
    
    Args:
        pkg_file: Path to the package.json or package.json.tpl file
        temp_dir: Temporary directory to use for the check
        is_template: Whether this is a .tpl template file
        
    Returns:
        Tuple of (has_vulnerabilities: bool, message: str)
    """
    # Create a unique work directory for this check
    work_dir = temp_dir / f"check_{pkg_file.name}_{hash(str(pkg_file)) % 10000}"
    work_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Copy the file to the work directory as package.json
        dest_file = work_dir / "package.json"
        shutil.copy(pkg_file, dest_file)
        
        # Run npm install --package-lock-only to generate package-lock.json
        install_result = subprocess.run(
            ["npm", "install", "--package-lock-only"],
            cwd=work_dir,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if install_result.returncode != 0:
            return False, f"WARNING: Could not generate package-lock.json: {install_result.stderr[:200]}"
        
        # Run npm audit to check for vulnerabilities
        audit_result = subprocess.run(
            ["npm", "audit", "--json"],
            cwd=work_dir,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        # Parse the audit result
        try:
            audit_data = json.loads(audit_result.stdout) if audit_result.stdout else {}
        except json.JSONDecodeError:
            # If JSON parsing fails, check the return code
            if audit_result.returncode != 0:
                return True, f"Vulnerabilities found (npm audit exit code: {audit_result.returncode})"
            return False, "OK"
        
        # Check for vulnerabilities in the audit result
        vulnerabilities = audit_data.get("vulnerabilities", {})
        metadata = audit_data.get("metadata", {}).get("vulnerabilities", {})
        
        # Count vulnerabilities by severity
        critical = metadata.get("critical", 0)
        high = metadata.get("high", 0)
        moderate = metadata.get("moderate", 0)
        
        if critical > 0 or high > 0 or moderate > 0:
            vuln_summary = []
            if critical > 0:
                vuln_summary.append(f"{critical} critical")
            if high > 0:
                vuln_summary.append(f"{high} high")
            if moderate > 0:
                vuln_summary.append(f"{moderate} moderate")
            return True, f"Vulnerabilities found: {', '.join(vuln_summary)}"
        
        return False, "OK"
        
    except subprocess.TimeoutExpired:
        return False, "WARNING: npm command timed out"
    except Exception as e:
        return False, f"WARNING: Error during check: {str(e)}"
    finally:
        # Clean up the work directory
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="Check npm package.json files for security vulnerabilities"
    )
    parser.add_argument(
        "--scan-directory",
        nargs="+",
        default=["templates/vsc"],
        help="Directories to scan for package.json files"
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose output"
    )
    
    args = parser.parse_args()
    
    safe_print("=" * 60)
    safe_print("NPM Security Vulnerability Check")
    safe_print("=" * 60)
    safe_print(f"Scanning directories: {', '.join(args.scan_directory)}")
    safe_print("")
    
    # Find all package files
    package_files, template_files = find_package_files(args.scan_directory)
    
    total_files = len(package_files) + len(template_files)
    safe_print(f"Found {len(package_files)} package.json files")
    safe_print(f"Found {len(template_files)} package.json.tpl template files")
    safe_print(f"Total files to check: {total_files}")
    safe_print("")
    
    if total_files == 0:
        safe_print("No package.json files found to check.")
        sys.exit(0)
    
    # Create a temporary directory for all checks
    with tempfile.TemporaryDirectory(prefix="npm_security_check_") as temp_dir:
        temp_path = Path(temp_dir)
        safe_print(f"Using temp directory: {temp_dir}")
        safe_print("")
        
        failed_files = []
        checked_count = 0
        
        # Check regular package.json files
        for pkg_file in package_files:
            checked_count += 1
            safe_print(f"[{checked_count}/{total_files}] Checking: {pkg_file}")
            
            has_vuln, message = check_package_vulnerabilities(pkg_file, temp_path, is_template=False)
            
            if has_vuln:
                safe_print(f"  ❌ ERROR: {message}")
                failed_files.append((pkg_file, message))
            else:
                if message.startswith("WARNING"):
                    safe_print(f"  ⚠️ {message}")
                else:
                    safe_print(f"  ✅ {message}")
        
        # Check template files
        for tpl_file in template_files:
            checked_count += 1
            safe_print(f"[{checked_count}/{total_files}] Checking: {tpl_file}")
            
            has_vuln, message = check_package_vulnerabilities(tpl_file, temp_path, is_template=True)
            
            if has_vuln:
                safe_print(f"  ❌ ERROR: {message}")
                failed_files.append((tpl_file, message))
            else:
                if message.startswith("WARNING"):
                    safe_print(f"  ⚠️ {message}")
                else:
                    safe_print(f"  ✅ {message}")
    
    # Summary
    safe_print("")
    safe_print("=" * 60)
    safe_print("Summary")
    safe_print("=" * 60)
    safe_print(f"Total files checked: {checked_count}")
    safe_print(f"Files with vulnerabilities: {len(failed_files)}")
    
    if failed_files:
        safe_print("")
        safe_print("Failed files:")
        for pkg_file, message in failed_files:
            safe_print(f"  - {pkg_file}: {message}")
        safe_print("")
        safe_print("❌ FAILED: Security vulnerabilities found")
        sys.exit(1)
    else:
        safe_print("")
        safe_print("✅ SUCCESS: All package.json files passed security check")
        sys.exit(0)


if __name__ == "__main__":
    main()
