"""
Data models for the Dropbox downloader.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class DownloadResult:
    """Result of a single download operation."""
    success: bool
    message: str
    file_path: Optional[Path] = None


@dataclass
class DownloadStats:
    """Track download statistics and failures."""
    total: int = 0
    completed: int = 0
    skipped: int = 0
    failed: list = field(default_factory=list)

    def add_completed(self) -> None:
        self.completed += 1

    def add_skipped(self) -> None:
        self.skipped += 1

    def add_failed(self, upc: str, url: str, error: str, row_data: Optional[dict] = None) -> None:
        self.failed.append({
            'upc': upc,
            'url': url,
            'error': str(error),
            'row_data': row_data
        })

    def print_summary(self) -> None:
        print("\n" + "=" * 60)
        print("DOWNLOAD SUMMARY")
        print("=" * 60)
        print(f"Total items:     {self.total}")
        print(f"Downloaded:      {self.completed}")
        print(f"Skipped:         {self.skipped}")
        print(f"Failed:          {len(self.failed)}")
        print("=" * 60)

        if self.failed:
            print("\nFAILED DOWNLOADS:")
            for item in self.failed:
                print(f"  UPC: {item['upc']}")
                print(f"  URL: {item['url']}")
                print(f"  Error: {item['error']}")
                print()
