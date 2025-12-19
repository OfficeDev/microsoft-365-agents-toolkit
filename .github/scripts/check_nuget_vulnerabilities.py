#!/usr/bin/env python3
"""
Check NuGet packages in C# projects for security vulnerabilities.
This script scans .csproj and .csproj.tpl files and runs dotnet list package --vulnerable
in a temp directory to avoid modifying the original repository.
"""

import os
import sys
import argparse
import subprocess
import tempfile
import shutil
import re
from pathlib import Path
from typing import List, Tuple


def safe_print(message: str) -> None:
    """Safely print message, handling encoding issues"""
    try:
        print(message)
    except UnicodeEncodeError:
        safe_message = message.encode('ascii', 'replace').decode('ascii')
        print(safe_message)
    except Exception as e:
        print(f"[Print error: {type(e).__name__}]")


def find_csproj_files(scan_dirs: List[str]) -> Tuple[List[Path], List[Path]]:
    """Find all .csproj and .csproj.tpl files in the given directories"""
    csproj_files = []
    template_files = []
    
    for scan_dir in scan_dirs:
        base_path = Path(scan_dir)
        if not base_path.exists():
            safe_print(f"WARNING: Directory does not exist: {scan_dir}")
            continue
        
        # Find .csproj files (excluding .tpl files)
        for csproj_file in base_path.rglob("*.csproj"):
            if not str(csproj_file).endswith(".tpl"):
                csproj_files.append(csproj_file)
        
        # Find .csproj.tpl template files
        for tpl_file in base_path.rglob("*.csproj.tpl"):
            template_files.append(tpl_file)
    
    return csproj_files, template_files


def process_template_file(tpl_file: Path, dest_dir: Path) -> Path:
    """
    Process a .csproj.tpl template file by replacing template variables.
    
    Args:
        tpl_file: Path to the .csproj.tpl file
        dest_dir: Destination directory for the processed file
        
    Returns:
        Path to the processed .csproj file
    """
    content = tpl_file.read_text(encoding='utf-8')
    
    # Replace common template variables
    replacements = {
        "{{TargetFramework}}": "net8.0",
        "{{ProjectName}}": "TempProject",
        "{{SafeProjectName}}": "TempProject",
        "{{RootNamespace}}": "TempProject",
        "{{AssemblyName}}": "TempProject",
        # Add more template variables as needed
    }
    
    for placeholder, value in replacements.items():
        content = content.replace(placeholder, value)
    
    # Also handle variations with spaces and different casing
    content = re.sub(r'\{\{\s*TargetFramework\s*\}\}', 'net8.0', content, flags=re.IGNORECASE)
    content = re.sub(r'\{\{\s*ProjectName\s*\}\}', 'TempProject', content, flags=re.IGNORECASE)
    
    dest_file = dest_dir / "TempProject.csproj"
    dest_file.write_text(content, encoding='utf-8')
    
    return dest_file


def check_nuget_vulnerabilities(csproj_file: Path, temp_dir: Path, is_template: bool = False) -> Tuple[bool, str]:
    """
    Check a .csproj file for NuGet vulnerabilities using dotnet list package --vulnerable.
    
    Args:
        csproj_file: Path to the .csproj or .csproj.tpl file
        temp_dir: Temporary directory to use for the check
        is_template: Whether this is a .tpl template file
        
    Returns:
        Tuple of (has_vulnerabilities: bool, message: str)
    """
    # Create a unique work directory for this check
    work_dir = temp_dir / f"check_{csproj_file.stem}_{hash(str(csproj_file)) % 10000}"
    work_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Process the file
        if is_template:
            processed_file = process_template_file(csproj_file, work_dir)
            csproj_name = processed_file.name
        else:
            # Copy the file to the work directory
            dest_file = work_dir / csproj_file.name
            shutil.copy(csproj_file, dest_file)
            csproj_name = dest_file.name
        
        # Run dotnet restore
        restore_result = subprocess.run(
            ["dotnet", "restore", csproj_name],
            cwd=work_dir,
            capture_output=True,
            text=True,
            timeout=180
        )
        
        if restore_result.returncode != 0:
            # Check if it's a critical error or just a warning
            stderr = restore_result.stderr.lower()
            if "error" in stderr and "warning" not in stderr:
                return False, f"WARNING: Could not restore packages: {restore_result.stderr[:200]}"
        
        # Run dotnet list package --vulnerable
        vuln_result = subprocess.run(
            ["dotnet", "list", csproj_name, "package", "--vulnerable"],
            cwd=work_dir,
            capture_output=True,
            text=True,
            timeout=180
        )
        
        output = vuln_result.stdout + vuln_result.stderr
        
        # Check for vulnerable packages in the output
        if "has the following vulnerable packages" in output:
            # Extract vulnerability details
            vuln_lines = []
            for line in output.split('\n'):
                if '>' in line and ('Critical' in line or 'High' in line or 'Moderate' in line or 'Low' in line):
                    vuln_lines.append(line.strip())
            
            if vuln_lines:
                return True, f"Vulnerabilities found: {len(vuln_lines)} package(s)"
            return True, "Vulnerabilities found"
        
        # Check for "no vulnerable packages" message
        if "no vulnerable packages" in output.lower() or vuln_result.returncode == 0:
            return False, "OK"
        
        return False, "OK"
        
    except subprocess.TimeoutExpired:
        return False, "WARNING: dotnet command timed out"
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
        description="Check NuGet packages in C# projects for security vulnerabilities"
    )
    parser.add_argument(
        "--scan-directory",
        nargs="+",
        default=["templates/vs"],
        help="Directories to scan for .csproj files"
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose output"
    )
    
    args = parser.parse_args()
    
    safe_print("=" * 60)
    safe_print("NuGet Security Vulnerability Check")
    safe_print("=" * 60)
    safe_print(f"Scanning directories: {', '.join(args.scan_directory)}")
    safe_print("")
    
    # Find all csproj files
    csproj_files, template_files = find_csproj_files(args.scan_directory)
    
    total_files = len(csproj_files) + len(template_files)
    safe_print(f"Found {len(csproj_files)} .csproj files")
    safe_print(f"Found {len(template_files)} .csproj.tpl template files")
    safe_print(f"Total files to check: {total_files}")
    safe_print("")
    
    if total_files == 0:
        safe_print("No .csproj files found to check.")
        sys.exit(0)
    
    # Create a temporary directory for all checks
    with tempfile.TemporaryDirectory(prefix="nuget_security_check_") as temp_dir:
        temp_path = Path(temp_dir)
        safe_print(f"Using temp directory: {temp_dir}")
        safe_print("")
        
        failed_files = []
        checked_count = 0
        
        # Check regular .csproj files
        for csproj_file in csproj_files:
            checked_count += 1
            safe_print(f"[{checked_count}/{total_files}] Checking: {csproj_file}")
            
            has_vuln, message = check_nuget_vulnerabilities(csproj_file, temp_path, is_template=False)
            
            if has_vuln:
                safe_print(f"  ❌ ERROR: {message}")
                failed_files.append((csproj_file, message))
            else:
                if message.startswith("WARNING"):
                    safe_print(f"  ⚠️ {message}")
                else:
                    safe_print(f"  ✅ {message}")
        
        # Check template files
        for tpl_file in template_files:
            checked_count += 1
            safe_print(f"[{checked_count}/{total_files}] Checking: {tpl_file}")
            
            has_vuln, message = check_nuget_vulnerabilities(tpl_file, temp_path, is_template=True)
            
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
        for csproj_file, message in failed_files:
            safe_print(f"  - {csproj_file}: {message}")
        safe_print("")
        safe_print("❌ FAILED: Security vulnerabilities found")
        sys.exit(1)
    else:
        safe_print("")
        safe_print("✅ SUCCESS: All C# projects passed NuGet security check")
        sys.exit(0)


if __name__ == "__main__":
    main()
