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

class ImprovedReadmeImageAnalyzer:
    def __init__(self, base_path: str, scan_patterns: List[str] = None, exclude_dirs: List[str] = None, 
                 max_concurrent: int = 5, request_timeout: int = 10, max_total_time: int = 600):
        self.base_path = Path(base_path)
        self.scan_patterns = scan_patterns or ["**/README.md", "**/README.md.tpl", "**/CHANGELOG.md", "**/PRERELEASE.md"]
        self.exclude_dirs = exclude_dirs or []
        self.max_concurrent = max_concurrent
        self.request_timeout = request_timeout
        self.max_total_time = max_total_time
        self.start_time = time.time()
        self.last_request_time = 0
        self.min_request_interval = 0.4  # Minimum 0.4s between requests
        # Normalize exclude directories to Path objects for easier comparison
        self.exclude_paths = [Path(d) for d in self.exclude_dirs]
        self.results = {
            "files_analyzed": [],
            "images_found": [],
            "broken_images": [],
            "working_images": [],
            "summary": {}
        }
        
        # Image link regex patterns
        self.image_patterns = [
            r'!\[.*?\]\((.*?)\)',  # Markdown format: ![alt](url)
            r'<img[^>]+src="([^"]+)"[^>]*>',  # HTML img tag
            r'<img[^>]+src=\'([^\']+)\'[^>]*>',  # HTML img tag (single quotes)
        ]
        # Hyperlink regex (exclude image links)
        self.url_pattern = re.compile(r'(?<!!)\[[^]]+\]\((https?://[^)]+)\)')
        self.allowed_hostnames = {"aka.ms", "github.com", "microsoft.com", "visualstudio.com", "githubusercontent.com"}
        # Session settings to avoid repeated connections
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def extract_links_from_content(self, content: str, file_path: Path) -> list:
        """Extract all hyperlinks from file content (excluding images)"""
        links = []
        for match in self.url_pattern.findall(content):
            url = match.strip()
            if url:
                links.append({
                    "url": url,
                    "file": str(file_path.relative_to(self.base_path)),
                    "type": "absolute"
                })
        return links

    def find_readme_files(self) -> List[Path]:
        """Find all README.md and README.md.tpl files, excluding specified directories (case-insensitive)"""
        readme_files = []
        
        # Use configured scan patterns to find files with case-insensitive matching
        for pattern in self.scan_patterns:
            readme_files.extend(self._glob_case_insensitive(pattern))
        
        # Filter out files in excluded directories
        if self.exclude_dirs:
            filtered_files = []
            for file_path in readme_files:
                should_exclude = False
                relative_path = file_path.relative_to(self.base_path)
                
                # Check if file is in any excluded directory
                for exclude_dir in self.exclude_dirs:
                    exclude_path = Path(exclude_dir)
                    try:
                        # Check if the file path starts with the excluded directory
                        relative_path.relative_to(exclude_path)
                        should_exclude = True
                        break
                    except ValueError:
                        # relative_to raises ValueError if the path is not relative to exclude_path
                        continue
                
                if not should_exclude:
                    filtered_files.append(file_path)
            
            readme_files = filtered_files
        
        return sorted(readme_files)
    
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

    def extract_images_from_content(self, content: str, file_path: Path) -> List[Dict]:
        """Extract all image links from file content"""
        images = []
        
        for pattern in self.image_patterns:
            matches = re.findall(pattern, content, re.IGNORECASE | re.DOTALL)
            for match in matches:
                # Clean URL
                image_url = match.strip()
                if image_url:
                    images.append({
                        "url": image_url,
                        "file": str(file_path.relative_to(self.base_path)),
                        "type": self._get_url_type(image_url, file_path)
                    })
        
        return images

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

    def check_image_availability(self, image_info: Dict) -> Dict:
        """Check availability of a single image - prioritize local file system"""
        url = image_info["url"]
        file_path = Path(self.base_path) / image_info["file"]
        
        result = {
            **image_info,
            "local_exists": None,
            "local_path": None,
            "http_status": None,
            "http_available": None,
            "status": "unknown",
            "error": None
        }
        
        try:
            # For local paths, prioritize checking file system
            if image_info["type"] in ["relative", "root_relative"]:
                local_path = self._resolve_local_path(url, file_path)
                result["local_path"] = str(local_path)
                result["local_exists"] = local_path.exists()
                
                if result["local_exists"]:
                    result["status"] = "working"
                    return result
                else:
                    result["status"] = "broken"
                    result["error"] = f"Local file does not exist: {local_path}"
                    return result
            
            # For remote URLs, perform HTTP check
            else:
                resolved_url = url if url.startswith(('http://', 'https://')) else f"https:{url}"
                result["resolved_url"] = resolved_url
                
                # Apply rate limiting
                self._rate_limit()
                
                # Check HTTP availability with retry logic for 429 errors
                max_retries = 6
                retry_delays = [2, 4, 8, 16, 32, 64]  # Exponential backoff
                
                for attempt in range(max_retries + 1):
                    try:
                        response = self.session.head(resolved_url, timeout=self.request_timeout, allow_redirects=True)
                        http_status = response.status_code
                        
                        # If HEAD request fails, try GET request (some servers don't support HEAD)
                        if http_status >= 400:
                            try:
                                response = self.session.get(resolved_url, timeout=self.request_timeout, allow_redirects=True, stream=True)
                                http_status = response.status_code
                                # Only read a small part of content to confirm this is an image
                                response.close()
                            except:
                                pass
                        
                        # Handle rate limiting
                        if http_status == 429 and attempt < max_retries:
                            retry_delay = retry_delays[attempt]
                            safe_print(f"Rate limited (429), retrying in {retry_delay}s... (attempt {attempt + 1}/{max_retries + 1})")
                            time.sleep(retry_delay)
                            continue
                        
                        # Break out of retry loop on success or non-429 error
                        break
                        
                    except Exception as e:
                        if attempt < max_retries:
                            retry_delay = retry_delays[attempt]
                            safe_print(f"Request failed ({str(e)}), retrying in {retry_delay}s... (attempt {attempt + 1}/{max_retries + 1})")
                            time.sleep(retry_delay)
                            continue
                        else:
                            # Last attempt failed
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
        """Analyze all README files for images (local/remote) and hyperlinks (text links, not images)."""
        readme_files = self.find_readme_files()
        print(f"Found {len(readme_files)} README files")

        all_local_images = []
        all_remote_images = []
        all_hyperlinks = []
        remote_image_url_seen = set()
        remote_link_url_seen = set()

        for file_path in readme_files:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()

                # Extract image links
                images = self.extract_images_from_content(content, file_path)
                # Extract hyperlinks (text links, not images)
                links = self.extract_links_from_content(content, file_path)

                # Classify images
                for img in images:
                    if img["type"] in ["relative", "root_relative"]:
                        all_local_images.append(img)
                    else:
                        if img["url"] not in remote_image_url_seen:
                            remote_image_url_seen.add(img["url"])
                            all_remote_images.append(img)

                # Classify hyperlinks (only text links, not images)
                for link in links:
                    hostname = urlparse(link["url"]).hostname
                    if hostname and any(hostname.endswith(allowed) for allowed in self.allowed_hostnames):
                        if link["url"] not in remote_link_url_seen:
                            remote_link_url_seen.add(link["url"])
                            all_hyperlinks.append(link)

                self.results["files_analyzed"].append({
                    "file": str(file_path.relative_to(self.base_path)),
                    "images_count": len(images),
                    "hyperlinks_count": len(links)
                })

                try:
                    safe_path = str(file_path.relative_to(self.base_path)).encode('ascii', 'replace').decode('ascii')
                    safe_print(f"Analyzing file: {safe_path} - found {len(images)} images, {len(links)} hyperlinks")
                except Exception:
                    safe_print(f"Analyzing file: [file with special chars] - found {len(images)} images, {len(links)} hyperlinks")

            except Exception as e:
                print(f"Failed to read file {file_path}: {e}")

        # IMAGE CHECK: local image file check
        print(f"\nFound {len(all_local_images)} local images, {len(all_remote_images)} unique remote images.")
        image_check_results = {"local": [], "remote": [], "broken": [], "working": []}
        for i, img in enumerate(all_local_images):
            result = self.check_image_availability(img)
            image_check_results["local"].append(result)
            if result["status"] == "working":
                image_check_results["working"].append(result)
            else:
                image_check_results["broken"].append(result)
            try:
                safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                safe_print(f"Local image check {i+1}/{len(all_local_images)}: {safe_url}... - {result['status']}")
            except Exception:
                safe_print(f"Local image check {i+1}/{len(all_local_images)}: [URL with special chars]... - {result['status']}")

        # IMAGE CHECK: remote image link check (concurrent)
        if all_remote_images:
            print(f"Starting to check remote image URLs (max {self.max_concurrent} concurrent, {self.request_timeout}s timeout each)...")
            print(f"Maximum total time allowed: {self.max_total_time}s")
            try:
                with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
                    future_to_item = {executor.submit(self.check_image_availability, item): item for item in all_remote_images}
                    completed_count = 0
                    for i, future in enumerate(as_completed(future_to_item)):
                        if self._is_time_exceeded():
                            print(f"\n⚠️  Time limit exceeded ({self.max_total_time}s). Stopping after {completed_count} remote image URLs.")
                            for remaining_future in future_to_item:
                                if not remaining_future.done():
                                    remaining_future.cancel()
                            break
                        try:
                            result = future.result()
                            image_check_results["remote"].append(result)
                            completed_count += 1
                            if result["status"] == "working":
                                image_check_results["working"].append(result)
                            else:
                                image_check_results["broken"].append(result)
                            try:
                                safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                                elapsed = int(time.time() - self.start_time)
                                safe_print(f"Remote image check {completed_count}/{len(all_remote_images)} ({elapsed}s): {safe_url}... - {result['status']}")
                            except Exception:
                                safe_print(f"Remote image check {completed_count}/{len(all_remote_images)} ({elapsed}s): [URL with special chars]... - {result['status']}")
                        except Exception as e:
                            safe_print(f"Error processing remote image URL {completed_count+1}/{len(all_remote_images)}: {str(e)}")
                            completed_count += 1
            except Exception as e:
                safe_print(f"Error in concurrent processing: {str(e)}")
                for i, item in enumerate(all_remote_images):
                    try:
                        result = self.check_image_availability(item)
                        image_check_results["remote"].append(result)
                        if result["status"] == "working":
                            image_check_results["working"].append(result)
                        else:
                            image_check_results["broken"].append(result)
                        safe_print(f"Remote image check (sequential) {i+1}/{len(all_remote_images)}: [URL]... - {result['status']}")
                    except Exception as e:
                        safe_print(f"Error in sequential processing {i+1}/{len(all_remote_images)}: {str(e)}")

        # HYPERLINK CHECK: only allowed_hostnames, deduped, remote check only
        print(f"\nFound {len(all_hyperlinks)} unique allowed remote hyperlinks.")
        hyperlink_check_results = {"remote": [], "broken": [], "working": []}
        if all_hyperlinks:
            print(f"Starting to check remote hyperlinks (max {self.max_concurrent} concurrent, {self.request_timeout}s timeout each)...")
            print(f"Maximum total time allowed: {self.max_total_time}s")
            try:
                with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
                    future_to_item = {executor.submit(self.check_image_availability, item): item for item in all_hyperlinks}
                    completed_count = 0
                    for i, future in enumerate(as_completed(future_to_item)):
                        if self._is_time_exceeded():
                            print(f"\n⚠️  Time limit exceeded ({self.max_total_time}s). Stopping after {completed_count} hyperlinks.")
                            for remaining_future in future_to_item:
                                if not remaining_future.done():
                                    remaining_future.cancel()
                            break
                        try:
                            result = future.result()
                            hyperlink_check_results["remote"].append(result)
                            completed_count += 1
                            if result["status"] == "working":
                                hyperlink_check_results["working"].append(result)
                            else:
                                hyperlink_check_results["broken"].append(result)
                            try:
                                safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                                elapsed = int(time.time() - self.start_time)
                                safe_print(f"Hyperlink check {completed_count}/{len(all_hyperlinks)} ({elapsed}s): {safe_url}... - {result['status']}")
                            except Exception:
                                safe_print(f"Hyperlink check {completed_count}/{len(all_hyperlinks)} ({elapsed}s): [URL with special chars]... - {result['status']}")
                        except Exception as e:
                            safe_print(f"Error processing hyperlink {completed_count+1}/{len(all_hyperlinks)}: {str(e)}")
                            completed_count += 1
            except Exception as e:
                safe_print(f"Error in concurrent processing: {str(e)}")
                for i, item in enumerate(all_hyperlinks):
                    try:
                        result = self.check_image_availability(item)
                        hyperlink_check_results["remote"].append(result)
                        if result["status"] == "working":
                            hyperlink_check_results["working"].append(result)
                        else:
                            hyperlink_check_results["broken"].append(result)
                        safe_print(f"Hyperlink check (sequential) {i+1}/{len(all_hyperlinks)}: [URL]... - {result['status']}")
                    except Exception as e:
                        safe_print(f"Error in sequential processing {i+1}/{len(all_hyperlinks)}: {str(e)}")

        # Save two separate reports
        with open(self.base_path / "image_check_report.json", 'w', encoding='utf-8') as f:
            json.dump(image_check_results, f, indent=2, ensure_ascii=False)
        with open(self.base_path / "hyperlink_check_report.json", 'w', encoding='utf-8') as f:
            json.dump(hyperlink_check_results, f, indent=2, ensure_ascii=False)

        # Generate summary including both image and hyperlink check results
        self.results["summary"] = {
            "total_files": len(readme_files),
            "total_images": len(all_local_images) + len(all_remote_images),
            "total_hyperlinks": len(all_hyperlinks),
            "local_items": len(all_local_images),
            "remote_image_items": len(all_remote_images),
            "remote_hyperlink_items": len(all_hyperlinks),
            "working_images": len(image_check_results["working"]),
            "broken_images": len(image_check_results["broken"]),
            "success_rate_images": len(image_check_results["working"]) / (len(all_local_images) + len(all_remote_images)) * 100 if (len(all_local_images) + len(all_remote_images)) else 0,
            "working_hyperlinks": len(hyperlink_check_results["working"]),
            "broken_hyperlinks": len(hyperlink_check_results["broken"]),
            "success_rate_hyperlinks": len(hyperlink_check_results["working"]) / len(all_hyperlinks) * 100 if all_hyperlinks else 0
        }

        return self.results

    def print_summary(self):
        """Print analysis summary"""
        summary = self.results["summary"]
        print("\n" + "="*60)
        print("Image Check Summary")
        print("="*60)
        print(f"README files analyzed: {summary['total_files']}")
        print(f"Total images found: {summary['total_images']}")
        print(f"  - Local images: {summary['local_items']}")
        print(f"  - Remote images: {summary['remote_image_items']}")
        print(f"Working images: {summary['working_images']}")
        print(f"Broken images: {summary['broken_images']}")
        # Use success_rate_images for image summary, fallback to 'success_rate' for backward compatibility
        if 'success_rate_images' in summary:
            print(f"Success rate: {summary['success_rate_images']:.1f}%")
        elif 'success_rate' in summary:
            print(f"Success rate: {summary['success_rate']:.1f}%")
        else:
            print("Success rate: N/A")

        if self.results["broken_images"]:
            print(f"\nBroken image links:")
            for img in self.results["broken_images"]:
                print(f"  File: {img['file']}")
                print(f"  Link: {img['url']}")
                if img.get('local_path'):
                    print(f"  Local path: {img['local_path']}")
                    print(f"  File exists: {img.get('local_exists', 'N/A')}")
                if img.get('http_status'):
                    print(f"  HTTP status: {img.get('http_status', 'N/A')}")
                print(f"  Error: {img.get('error', 'N/A')}")
                print(f"---------------------------------------------")
                print()

        # Print hyperlink check summary
        print("\n" + "-"*60)
        print("Hyperlink Check Summary")
        print("-"*60)
        print(f"Total hyperlinks checked: {summary['total_hyperlinks']}")
        print(f"  - Remote hyperlinks (allowed hostnames): {summary['remote_hyperlink_items']}")
        print(f"Working hyperlinks: {summary['working_hyperlinks']}")
        print(f"Broken hyperlinks: {summary['broken_hyperlinks']}")
        if 'success_rate_hyperlinks' in summary:
            print(f"Success rate: {summary['success_rate_hyperlinks']:.1f}%")
        else:
            print("Success rate: N/A")
        if summary.get('broken_hyperlinks', 0) > 0:
            print(f"\nBroken hyperlinks:")
            # Optionally, could print details if needed (not implemented here)

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
    analyzer = ImprovedReadmeImageAnalyzer(
        scan_directory, 
        args.file_patterns, 
        args.exclude_dirs,
        args.max_concurrent,
        args.request_timeout,
        args.max_total_time
    )
    
    # Execute analysis
    results = analyzer.analyze_files()
    
    # Print summary
    analyzer.print_summary()
        
    # Check if all image links are available
    summary = results["summary"]
    success_rate = summary.get("success_rate_images", summary.get("success_rate", 0))
    broken_count = summary.get("broken_images", 0)

    print(f"\n{'='*60}")
    print("PIPELINE CHECK RESULTS")
    print(f"{'='*60}")

    if broken_count == 0:
        print("✅ All image links are working!")
        print("✅ PIPELINE check passed - can continue execution")
        exit_code = 0
    else:
        print(f"❌ Found {broken_count} broken image links!")
        print("❌ PIPELINE check failed - recommend stopping execution")
        exit_code = 1

    print(f"Exit code: {exit_code}")

    # Return result for Pipeline use
    return exit_code

if __name__ == "__main__":
    import sys
    exit_code = main()
    sys.exit(exit_code)