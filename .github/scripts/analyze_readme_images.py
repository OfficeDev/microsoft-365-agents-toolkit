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
    def __init__(self, base_path: str, scan_patterns: List[str] = None):
        self.base_path = Path(base_path)
        self.scan_patterns = scan_patterns or ["**/README.md", "**/README.md.tpl"]
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
        
        # Session settings to avoid repeated connections
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def find_readme_files(self) -> List[Path]:
        """Find all README.md and README.md.tpl files"""
        readme_files = []
        
        # Use configured scan patterns to find files
        for pattern in self.scan_patterns:
            readme_files.extend(self.base_path.glob(pattern))
        
        return sorted(readme_files)

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
                
                # Check HTTP availability
                response = self.session.head(resolved_url, timeout=10, allow_redirects=True)
                http_status = response.status_code
                
                # If HEAD request fails, try GET request (some servers don't support HEAD)
                if http_status >= 400:
                    try:
                        response = self.session.get(resolved_url, timeout=10, allow_redirects=True, stream=True)
                        http_status = response.status_code
                        # Only read a small part of content to confirm this is an image
                        response.close()
                    except:
                        pass
                
                result["http_status"] = http_status
                result["http_available"] = 200 <= http_status < 400
                result["status"] = "working" if result["http_available"] else "broken"
                
                if not result["http_available"]:
                    result["error"] = f"HTTP error: {http_status}"
                
                return result
                
        except Exception as e:
            result["status"] = "error"
            result["error"] = str(e)
            return result

    def analyze_files(self) -> Dict:
        """Analyze all README files"""
        readme_files = self.find_readme_files()
        print(f"Found {len(readme_files)} README files")
        
        all_images = []
        
        # Analyze files one by one
        for file_path in readme_files:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # Extract image links
                images = self.extract_images_from_content(content, file_path)
                all_images.extend(images)
                
                self.results["files_analyzed"].append({
                    "file": str(file_path.relative_to(self.base_path)),
                    "images_count": len(images)
                })
                
                try:
                    safe_path = str(file_path.relative_to(self.base_path)).encode('ascii', 'replace').decode('ascii')
                    safe_print(f"Analyzing file: {safe_path} - found {len(images)} images")
                except Exception:
                    safe_print(f"Analyzing file: [file with special chars] - found {len(images)} images")
                
            except Exception as e:
                print(f"Failed to read file {file_path}: {e}")
        
        print(f"\nFound {len(all_images)} image links in total")
        print("Starting to check image availability...")
        
        # Check image availability - check local files directly, remote files with concurrent requests
        local_images = [img for img in all_images if img["type"] in ["relative", "root_relative"]]
        remote_images = [img for img in all_images if img["type"] in ["absolute", "protocol_relative"]]
        
        print(f"Local images: {len(local_images)}, Remote images: {len(remote_images)}")
        
        # Process local images
        for i, img in enumerate(local_images):
            result = self.check_image_availability(img)
            self.results["images_found"].append(result)
            
            if result["status"] == "working":
                self.results["working_images"].append(result)
            else:
                self.results["broken_images"].append(result)
            
            try:
                safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                safe_print(f"Local check {i+1}/{len(local_images)}: {safe_url}... - {result['status']}")
            except Exception:
                safe_print(f"Local check {i+1}/{len(local_images)}: [URL with special chars]... - {result['status']}")
        
        # Process remote images concurrently
        if remote_images:
            print("Starting to check remote images...")
            try:
                with ThreadPoolExecutor(max_workers=10) as executor:
                    future_to_image = {executor.submit(self.check_image_availability, img): img for img in remote_images}
                    
                    for i, future in enumerate(as_completed(future_to_image)):
                        try:
                            result = future.result()
                            self.results["images_found"].append(result)
                            
                            if result["status"] == "working":
                                self.results["working_images"].append(result)
                            else:
                                self.results["broken_images"].append(result)
                            
                            try:
                                safe_url = result['url'][:50].encode('ascii', 'replace').decode('ascii')
                                safe_print(f"Remote check {i+1}/{len(remote_images)}: {safe_url}... - {result['status']}")
                            except Exception:
                                safe_print(f"Remote check {i+1}/{len(remote_images)}: [URL with special chars]... - {result['status']}")
                            
                            # Avoid too frequent requests
                            time.sleep(0.1)
                        except Exception as e:
                            safe_print(f"Error processing remote image {i+1}/{len(remote_images)}: {str(e)}")
            except Exception as e:
                safe_print(f"Error in concurrent processing: {str(e)}")
                # Fallback to sequential processing
                for i, img in enumerate(remote_images):
                    try:
                        result = self.check_image_availability(img)
                        self.results["images_found"].append(result)
                        
                        if result["status"] == "working":
                            self.results["working_images"].append(result)
                        else:
                            self.results["broken_images"].append(result)
                        
                        safe_print(f"Remote check (sequential) {i+1}/{len(remote_images)}: [URL]... - {result['status']}")
                    except Exception as e:
                        safe_print(f"Error in sequential processing {i+1}/{len(remote_images)}: {str(e)}")
        
        # Generate summary
        self.results["summary"] = {
            "total_files": len(readme_files),
            "total_images": len(all_images),
            "local_images": len(local_images),
            "remote_images": len(remote_images),
            "working_images": len(self.results["working_images"]),
            "broken_images": len(self.results["broken_images"]),
            "success_rate": len(self.results["working_images"]) / len(all_images) * 100 if all_images else 0
        }
        
        return self.results

    def save_results(self, output_file: str = "improved_readme_image_analysis.json"):
        """Save analysis results to JSON file"""
        output_path = self.base_path / output_file
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        print(f"Results saved to: {output_path}")

    def print_summary(self):
        """Print analysis summary"""
        summary = self.results["summary"]
        print("\n" + "="*60)
        print("Improved Image Link Analysis Summary")
        print("="*60)
        print(f"README files analyzed: {summary['total_files']}")
        print(f"Total images found: {summary['total_images']}")
        print(f"  - Local images: {summary['local_images']}")
        print(f"  - Remote images: {summary['remote_images']}")
        print(f"Working images: {summary['working_images']}")
        print(f"Broken images: {summary['broken_images']}")
        print(f"Success rate: {summary['success_rate']:.1f}%")
        
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
    
    args = parser.parse_args()
    
    # Determine scan directory
    scan_directory = args.scan_directory if args.scan_directory else os.getcwd()
    
    print(f"Scan directory: {scan_directory}")
    print(f"File patterns: {', '.join(args.file_patterns)}")
    
    # Create analyzer instance
    analyzer = ImprovedReadmeImageAnalyzer(scan_directory, args.file_patterns)
    
    # Execute analysis
    results = analyzer.analyze_files()
    
    # Print summary
    analyzer.print_summary()
    
    # Save results
    analyzer.save_results()
    
    # Check if all image links are available
    success_rate = results["summary"]["success_rate"]
    broken_count = results["summary"]["broken_images"]
    
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