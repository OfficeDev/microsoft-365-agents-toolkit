#!/usr/bin/env python3
"""
Improved version: Prioritize checking local file system, only perform HTTP checks for remote links
Support custom scan directories and file types
"""

import os
import re
import requests
import time
import argparse
import fnmatch
from pathlib import Path
from urllib.parse import urljoin, urlparse
from typing import List, Dict, Tuple
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys

def safe_print(message):
    """Safely print message, handling encoding issues"""
    try:
        print(message)
    except UnicodeEncodeError:
        # Fallback to ASCII with replacement characters
        safe_message = message.encode('ascii', 'replace').decode('ascii')
        print(safe_message)
    except Exception as e:
        # Ultimate fallback
        print(f"[Print error: {type(e).__name__}]")

class MarkdownUrlLinkChecker:
    def __init__(self, base_path: str, scan_patterns: List[str] = None, exclude_dirs: List[str] = None,
                 max_concurrent: int = 2, request_timeout: int = 10, max_total_time: int = 600):
        self.base_path = Path(base_path)
        self.scan_patterns = scan_patterns or ["**/README.md", "**/README.md.tpl", "**/CHANGELOG.md", "**/PRERELEASE.md"]
        self.exclude_dirs = exclude_dirs or []
        self.max_concurrent = max_concurrent
        self.request_timeout = request_timeout
        self.max_total_time = max_total_time
        self.start_time = time.time()
        self.last_request_time = 0
        self.min_request_interval = 1.0  # 1s between requests
        self.results = {
            "files_analyzed": [],
            "links_found": [],
            "broken_links": [],
            "working_links": [],
            "summary": {}
        }
        # Markdown URL link regex (not images)
        self.url_pattern = re.compile(r'(?<!\!)\[[^\]]+\]\((https?://[^)]+)\)')
        self.allowed_hostnames = {"aka.ms", "github.com", "microsoft.com", "visualstudio.com"}
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def find_files(self) -> List[Path]:
        files = []
        for pattern in self.scan_patterns:
            files.extend(self._glob_case_insensitive(pattern))
        if self.exclude_dirs:
            filtered_files = []
            for file_path in files:
                should_exclude = False
                relative_path = file_path.relative_to(self.base_path)
                for exclude_dir in self.exclude_dirs:
                    exclude_path = Path(exclude_dir)
                    try:
                        relative_path.relative_to(exclude_path)
                        should_exclude = True
                        break
                    except ValueError:
                        continue
                if not should_exclude:
                    filtered_files.append(file_path)
            files = filtered_files
        return sorted(files)
    
    def _glob_case_insensitive(self, pattern: str) -> List[Path]:
        """Perform case-insensitive glob matching"""
        matching_files = []
        
        # Handle different pattern types
        if pattern.startswith('**/'):  # Recursive pattern
            pattern_without_recursive = pattern[3:]  # Remove '**/' prefix
            for root, dirs, files in os.walk(self.base_path):
                root_path = Path(root)
                for file in files:
                    if fnmatch.fnmatch(file.lower(), pattern_without_recursive.lower()):
                        matching_files.append(root_path / file)
        else:
            # Non-recursive pattern - use regular glob but check case-insensitively
            all_files = self.base_path.glob(pattern.replace('*', '*').replace('?', '?'))
            for file_path in all_files:
                if fnmatch.fnmatch(file_path.name.lower(), Path(pattern).name.lower()):
                    matching_files.append(file_path)
        
        return matching_files
    
    def _is_time_exceeded(self) -> bool:
        """Check if maximum execution time has been exceeded"""
        return (time.time() - self.start_time) > self.max_total_time
    
    def _rate_limit(self):
        """Enforce rate limiting between requests"""
        current_time = time.time()
        time_since_last = current_time - self.last_request_time
        if time_since_last < self.min_request_interval:
            sleep_time = self.min_request_interval - time_since_last
            time.sleep(sleep_time)
        self.last_request_time = time.time()

    def extract_links_from_content(self, content: str, file_path: Path) -> List[Dict]:
        links = []
        for match in self.url_pattern.findall(content):
            url = match.strip()
            if url:
                # Only check allowed hostnames
                try:
                    hostname = url.split("//", 1)[-1].split("/", 1)[0].lower()
                    # Check for subdomains as well
                    if any(hostname == allowed or hostname.endswith('.' + allowed) for allowed in self.allowed_hostnames):
                        links.append({
                            "url": url,
                            "file": str(file_path.relative_to(self.base_path))
                        })
                except Exception:
                    continue
        return links

    def _get_url_type(self, url: str, file_path: Path) -> str:
        """Determine URL type"""
        if url.startswith(('http://', 'https://')):
            return "absolute"
        elif url.startswith('//'):
            return "protocol_relative"
        elif url.startswith('/'):
            return "root_relative"
        else:
            return "relative"

    def _resolve_local_path(self, url: str, file_path: Path) -> Path:
        """Resolve relative URL to local absolute path"""
        if url.startswith('/'):
            # Root relative path
            return self.base_path / url.lstrip('/')
        else:
            # Relative path, based on current file location
            return file_path.parent / url

    def check_link_availability(self, link_info: Dict) -> Dict:
        url = link_info["url"]
        result = {
            **link_info,
            "http_status": None,
            "http_available": None,
            "status": "unknown",
            "error": None
        }
        max_retries = 6
        retry_delays = [2, 4, 8, 16, 32, 64]
        try:
            self._rate_limit()
            for attempt in range(max_retries + 1):
                try:
                    response = self.session.head(url, timeout=self.request_timeout, allow_redirects=True)
                    http_status = response.status_code
                    if http_status >= 400:
                        try:
                            response = self.session.get(url, timeout=self.request_timeout, allow_redirects=True, stream=True)
                            http_status = response.status_code
                            response.close()
                        except:
                            pass
                    if http_status == 429 and attempt < max_retries:
                        retry_delay = retry_delays[attempt]
                        safe_print(f"Rate limited (429), retrying in {retry_delay}s... (attempt {attempt + 1}/{max_retries + 1})")
                        time.sleep(retry_delay)
                        continue
                    break
                except Exception as e:
                    if attempt < max_retries:
                        retry_delay = retry_delays[attempt]
                        safe_print(f"Request failed ({str(e)}), retrying in {retry_delay}s... (attempt {attempt + 1}/{max_retries + 1})")
                        time.sleep(retry_delay)
                        continue
                    else:
                        http_status = 0
                        break
            result["http_status"] = http_status
            result["http_available"] = 200 <= http_status < 400
            result["status"] = "working" if result["http_available"] else "broken"
            if not result["http_available"]:
                if http_status == 429:
                    result["error"] = f"HTTP error: {http_status} (Rate Limited - too many requests)"
                else:
                    result["error"] = f"HTTP error: {http_status}"
            return result
        except Exception as e:
            result["status"] = "error"
            result["error"] = str(e)
            return result

    def analyze_files(self) -> Dict:
        files = self.find_files()
        print(f"Found {len(files)} files to check")
        all_links = []
        for file_path in files:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                links = self.extract_links_from_content(content, file_path)
                all_links.extend(links)
                self.results["files_analyzed"].append({
                    "file": str(file_path.relative_to(self.base_path)),
                    "links_count": len(links)
                })
                safe_print(f"Analyzing file: {file_path.relative_to(self.base_path)} - found {len(links)} links")
            except Exception as e:
                print(f"Failed to read file {file_path}: {e}")
        print(f"\nFound {len(all_links)} URL links in total")
        print("Starting to check link availability...")
        with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
            future_to_link = {executor.submit(self.check_link_availability, link): link for link in all_links}
            completed_count = 0
            for i, future in enumerate(as_completed(future_to_link)):
                if (time.time() - self.start_time) > self.max_total_time:
                    print(f"\n⚠️  Time limit exceeded ({self.max_total_time}s). Stopping after {completed_count} links.")
                    for remaining_future in future_to_link:
                        if not remaining_future.done():
                            remaining_future.cancel()
                    break
                try:
                    result = future.result()
                    self.results["links_found"].append(result)
                    completed_count += 1
                    if result["status"] == "working":
                        self.results["working_links"].append(result)
                    else:
                        self.results["broken_links"].append(result)
                    safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                    elapsed = int(time.time() - self.start_time)
                    safe_print(f"Link check {completed_count}/{len(all_links)} ({elapsed}s): {safe_url}... - {result['status']}")
                except Exception as e:
                    safe_print(f"Error processing link {completed_count+1}/{len(all_links)}: {str(e)}")
                    completed_count += 1
        self.results["summary"] = {
            "total_files": len(files),
            "total_links": len(all_links),
            "working_links": len(self.results["working_links"]),
            "broken_links": len(self.results["broken_links"]),
            "success_rate": len(self.results["working_links"]) / len(all_links) * 100 if all_links else 0
        }
        return self.results

    def save_results(self, output_file: str = "improved_readme_image_analysis.json"):
        """Save analysis results to JSON file"""
        output_path = self.base_path / output_file
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        print(f"Results saved to: {output_path}")

    def print_summary(self):
        summary = self.results["summary"]
        print("\n" + "="*60)
        print("Markdown URL Link Analysis Summary")
        print("="*60)
        print(f"Files analyzed: {summary['total_files']}")
        print(f"Total links found: {summary['total_links']}")
        print(f"Working links: {summary['working_links']}")
        print(f"Broken links: {summary['broken_links']}")
        print(f"Success rate: {summary['success_rate']:.1f}%")
        if self.results["broken_links"]:
            print(f"\nBroken links:")
            for link in self.results["broken_links"]:
                print(f"  File: {link['file']}")
                print(f"  Link: {link['url']}")
                print(f"  HTTP status: {link.get('http_status', 'N/A')}")
                print(f"  Error: {link.get('error', 'N/A')}")
                print(f"---------------------------------------------")
                print()

def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description="Analyze image link availability in README files")
    parser.add_argument(
        "--scan-directory", 
        "-d", 
        default=None,
        help="Directory path to scan (default: current working directory)"
    )
    parser.add_argument(
        "--file-patterns", 
        "-p", 
        nargs="+",
        default=["**/README.md", "**/README.md.tpl"],
        help="File patterns to scan (default: README.md and README.md.tpl)"
    )
    parser.add_argument(
        "--exclude-dirs",
        "-e",
        nargs="+",
        default=[],
        help="Directory patterns to exclude from scanning (e.g., node_modules, .git, temp)"
    )
    parser.add_argument(
        "--max-concurrent",
        "-c",
        type=int,
        default=2,
        help="Maximum number of concurrent HTTP requests (default: 2)"
    )
    parser.add_argument(
        "--request-timeout",
        "-t",
        type=int,
        default=10,
        help="HTTP request timeout in seconds (default: 10)"
    )
    parser.add_argument(
        "--max-total-time",
        "-m",
        type=int,
        default=600,
        help="Maximum total execution time in seconds (default: 600 = 10 minutes)"
    )
    
    args = parser.parse_args()
    
    # Determine scan directory
    scan_directory = args.scan_directory if args.scan_directory else os.getcwd()
    
    print(f"Scan directory: {scan_directory}")
    print(f"File patterns: {', '.join(args.file_patterns)}")
    if args.exclude_dirs:
        print(f"Excluded directories: {', '.join(args.exclude_dirs)}")
    print(f"Max concurrent requests: {args.max_concurrent}")
    print(f"Request timeout: {args.request_timeout}s")
    print(f"Max total time: {args.max_total_time}s")
    
    # Create analyzer instance
    checker = MarkdownUrlLinkChecker(
        scan_directory,
        args.file_patterns,
        args.exclude_dirs,
        args.max_concurrent,
        args.request_timeout,
        args.max_total_time
    )
    results = checker.analyze_files()
    checker.print_summary()
    broken_count = results["summary"]["broken_links"]
    print(f"\n{'='*60}")
    print("PIPELINE CHECK RESULTS")
    print(f"{'='*60}")
    if broken_count == 0:
        print("✅ All URL links are working!")
        print("✅ PIPELINE check passed - can continue execution")
        exit_code = 0
    else:
        print(f"❌ Found {broken_count} broken URL links!")
        print("❌ PIPELINE check failed - recommend stopping execution")
        exit_code = 1
    print(f"Exit code: {exit_code}")
    return exit_code

if __name__ == "__main__":
    import sys
    exit_code = main()
    sys.exit(exit_code)