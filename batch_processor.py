"""
Batch processor for downloading files from Dropbox shared folders.

Handles Excel file processing, parallel downloads, and failure tracking.
"""

import shutil
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Optional, Set, Tuple

import pandas as pd
from tqdm import tqdm

from dropbox_client import download_first_file
from models import DownloadResult, DownloadStats


def check_existing_file(output_dir: str, upc: str, category: Optional[str] = None) -> Optional[Path]:
    """
    Check if a file with the given UPC already exists in the output directory.
    
    Args:
        output_dir: Base output directory
        upc: UPC code to check for
        category: Optional category subdirectory
        
    Returns:
        Path to existing file if found, None otherwise
    """
    if category:
        output_path = Path(output_dir) / category
    else:
        output_path = Path(output_dir)

    if not output_path.exists():
        return None

    for file in output_path.iterdir():
        if file.is_file() and file.stem == str(upc):
            return file

    return None


def download_and_rename(
    upc: str,
    image_url: str,
    output_dir: str,
    debug: bool = False,
    thread_id: int = 0,
    progress_bar=None,
    category: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Download the first image from a Dropbox folder and rename it with the UPC.

    Args:
        upc: UPC code to use as filename
        image_url: Dropbox shared folder URL
        output_dir: Directory to save the file
        debug: Enable debug output
        thread_id: Thread identifier for unique Chrome profiles
        progress_bar: Optional tqdm progress bar instance
        category: Optional category to organize files into subdirectories

    Returns:
        Tuple of (success, message)
    """
    try:
        # Determine the target directory
        if category:
            target_dir = Path(output_dir) / category
        else:
            target_dir = Path(output_dir)
        target_dir.mkdir(parents=True, exist_ok=True)

        # Check if file already exists
        existing = check_existing_file(output_dir, upc, category)
        if existing:
            return (True, f"Skipped (already exists: {existing.name})")

        # Create a unique temp directory for this download
        temp_dir = Path(output_dir) / f".tmp_{thread_id}_{upc}"
        temp_dir.mkdir(parents=True, exist_ok=True)

        # Use a unique Chrome profile per thread
        import platform
        import tempfile
        if platform.system() == "Windows":
            base_temp = tempfile.gettempdir()
            user_data_dir = f"{base_temp}\\chrome-download-{thread_id}-{upc}"
        else:
            user_data_dir = f"/tmp/chrome-download-{thread_id}"

        try:
            downloaded_file = download_first_file(
                url=image_url,
                output_dir=str(temp_dir),
                debug=debug,
                use_alt_method=False,
                user_data_dir=user_data_dir,
                progress_bar=progress_bar,
                file_label=str(upc)
            )

            if not downloaded_file or not downloaded_file.exists():
                return (False, "Download failed - no file returned")

            # Rename and move to target directory
            extension = downloaded_file.suffix
            final_path = target_dir / f"{upc}{extension}"
            shutil.move(str(downloaded_file), str(final_path))

            # Clean up
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            user_data_path = Path(user_data_dir)
            if user_data_path.exists():
                shutil.rmtree(user_data_path, ignore_errors=True)

            return (True, f"Downloaded as {final_path.name}")

        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
            user_data_path = Path(user_data_dir)
            if user_data_path.exists():
                shutil.rmtree(user_data_path, ignore_errors=True)

    except Exception as e:
        return (False, f"Error: {str(e)}")


def create_failed_excel(df_failed: pd.DataFrame, output_dir: str, excel_file: str) -> Optional[Path]:
    """
    Create an Excel file with failed downloads.

    Args:
        df_failed: DataFrame with failed download rows
        output_dir: Directory where files are saved
        excel_file: Original Excel filename

    Returns:
        Path to the created failed Excel file, or None if no failures
    """
    if df_failed.empty:
        return None

    output_path = Path(output_dir)
    dir_name = output_path.name if output_path.name else output_path.parts[-1]

    failed_excel_path = Path.cwd() / f"failed_{dir_name}.xlsx"
    df_failed.to_excel(failed_excel_path, index=False)

    return failed_excel_path


def remove_successful_from_failed_excel(failed_excel_path: Path, successful_upcs: Set[str]) -> None:
    """
    Remove successfully downloaded entries from the failed Excel file.

    Args:
        failed_excel_path: Path to the failed Excel file
        successful_upcs: Set of UPCs that were successfully downloaded
    """
    if not failed_excel_path.exists():
        return

    try:
        df = pd.read_excel(failed_excel_path)
        df_remaining = df[~df['UPC'].astype(str).str.strip().isin(successful_upcs)]

        if df_remaining.empty:
            failed_excel_path.unlink()
            print(f"\n✓ All failed items successfully downloaded. Removed {failed_excel_path.name}")
        else:
            df_remaining.to_excel(failed_excel_path, index=False)
            print(f"\n✓ Updated {failed_excel_path.name} - {len(successful_upcs)} items removed, {len(df_remaining)} remaining")
    except Exception as e:
        print(f"\n⚠ Warning: Could not update failed Excel file: {e}")


def process_excel(
    excel_file: str,
    output_dir: str,
    threads: int = 1,
    debug: bool = False,
    no_categories: bool = False
) -> Optional[Path]:
    """
    Process Excel file and download images.

    Args:
        excel_file: Path to Excel file with UPC and IMAGES LINK columns
        output_dir: Directory to save downloaded files
        threads: Number of parallel download threads
        debug: Enable debug output
        no_categories: Ignore CATEGORY column if present

    Returns:
        Path to failed Excel file if there were failures, None otherwise
    """
    print(f"Reading Excel file: {excel_file}")
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        sys.exit(1)

    # Check if this is a retry
    excel_path = Path(excel_file)
    is_retry = excel_path.name.startswith('failed_')
    if is_retry:
        print("📝 Retrying failed downloads...")

    # Validate columns
    required_cols = ['UPC', 'IMAGES LINK']
    if not all(col in df.columns for col in required_cols):
        print("✗ Error: Excel file must contain 'UPC' and 'IMAGES LINK' columns")
        print(f"  Found columns: {', '.join(df.columns)}")
        sys.exit(1)

    has_category = 'CATEGORY' in df.columns and not no_categories
    if has_category:
        print("✓ CATEGORY column found - files will be organized by category")
    elif 'CATEGORY' in df.columns and no_categories:
        print("⊘ CATEGORY column found but ignored (--no-categories flag set)")

    # Filter out rows with missing data
    df = df.dropna(subset=['UPC', 'IMAGES LINK'])

    # Create output directory
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    stats = DownloadStats()
    stats.total = len(df)
    successful_upcs: Set[str] = set()

    print(f"Found {stats.total} items to process")
    print(f"Output directory: {output_path.resolve()}")
    print(f"Threads: {threads}")
    print()

    if threads == 1:
        _process_single_threaded(df, output_dir, debug, has_category, stats, successful_upcs)
    else:
        _process_multi_threaded(df, output_dir, debug, has_category, stats, successful_upcs, threads)

    stats.print_summary()

    # Handle failed downloads
    failed_excel_path = None
    if stats.failed:
        failed_rows = [item['row_data'] for item in stats.failed]
        df_failed = pd.DataFrame(failed_rows)
        failed_excel_path = create_failed_excel(df_failed, output_dir, excel_file)

        if failed_excel_path:
            print(f"\n📋 Failed downloads saved to: {failed_excel_path}")
            print(f"\n💡 To retry failed downloads only, run:")
            print(f"   python cli.py {failed_excel_path.name} {output_dir}")
            if threads > 1:
                print(f"   python cli.py {failed_excel_path.name} {output_dir} --threads {threads}")

    if is_retry and successful_upcs:
        remove_successful_from_failed_excel(excel_path, successful_upcs)

    return failed_excel_path if stats.failed else None


def _process_single_threaded(
    df: pd.DataFrame,
    output_dir: str,
    debug: bool,
    has_category: bool,
    stats: DownloadStats,
    successful_upcs: Set[str]
) -> None:
    """Process downloads in a single thread with progress bar."""
    with tqdm(total=stats.total, desc="Processing", unit="file", position=0) as pbar:
        for idx, row in df.iterrows():
            upc = str(row['UPC']).strip()
            url = str(row['IMAGES LINK']).strip()
            category = str(row['CATEGORY']).strip() if has_category and pd.notna(row.get('CATEGORY')) else None

            pbar.set_description(f"Processing {upc}")

            success, message = download_and_rename(
                upc, url, output_dir, debug, thread_id=0, progress_bar=pbar, category=category
            )

            if success:
                if "Skipped" in message:
                    stats.add_skipped()
                    pbar.write(f"⊘ {upc}: {message}")
                else:
                    stats.add_completed()
                    successful_upcs.add(upc)
                    pbar.write(f"✓ {upc}: {message}")
            else:
                stats.add_failed(upc, url, message, row_data=row.to_dict())
                pbar.write(f"✗ {upc}: {message}")

            pbar.update(1)


def _process_multi_threaded(
    df: pd.DataFrame,
    output_dir: str,
    debug: bool,
    has_category: bool,
    stats: DownloadStats,
    successful_upcs: Set[str],
    threads: int
) -> None:
    """Process downloads in multiple threads."""
    with tqdm(total=stats.total, desc="Overall Progress", unit="file") as overall_pbar:
        with ThreadPoolExecutor(max_workers=threads) as executor:
            future_to_item = {}
            for idx, row in df.iterrows():
                upc = str(row['UPC']).strip()
                url = str(row['IMAGES LINK']).strip()
                category = str(row['CATEGORY']).strip() if has_category and pd.notna(row.get('CATEGORY')) else None

                thread_id = idx % threads

                future = executor.submit(
                    download_and_rename,
                    upc, url, output_dir, debug,
                    thread_id=thread_id,
                    progress_bar=None,
                    category=category
                )
                future_to_item[future] = (idx, upc, url)

            for future in as_completed(future_to_item):
                idx, upc, url = future_to_item[future]

                try:
                    success, message = future.result()
                    row_data = df.loc[idx].to_dict()

                    if success:
                        if "Skipped" in message:
                            stats.add_skipped()
                            overall_pbar.write(f"⊘ {upc}: {message}")
                        else:
                            stats.add_completed()
                            successful_upcs.add(upc)
                            overall_pbar.write(f"✓ {upc}: {message}")
                    else:
                        stats.add_failed(upc, url, message, row_data=row_data)
                        overall_pbar.write(f"✗ {upc}: {message}")

                except Exception as e:
                    row_data = df.loc[idx].to_dict()
                    stats.add_failed(upc, url, str(e), row_data=row_data)
                    overall_pbar.write(f"✗ {upc}: Exception: {str(e)}")

                overall_pbar.update(1)
