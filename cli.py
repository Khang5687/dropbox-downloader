#!/usr/bin/env python3
"""
Command-line interface for Dropbox batch downloader.

Usage:
    python cli.py <excel_file> <output_dir> [options]
    
Examples:
    python cli.py products.xlsx output/
    python cli.py products.xlsx output/ --threads 4
    python cli.py products.xlsx output/ --retry
"""

import argparse
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Batch download images from Dropbox shared folders using Excel file input',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python cli.py products.xlsx output/
  python cli.py products.xlsx output/ --threads 4
  python cli.py products.xlsx output/ --retry
  python cli.py products.xlsx output/ --retry 3

Excel file format:
  Required columns:
    - UPC: Product UPC code (used as filename)
    - IMAGES LINK: Dropbox shared folder URL
  Optional columns:
    - CATEGORY: Organize files into subdirectories

Notes:
  - Files are saved as <UPC>.<extension> in the output directory
  - Failed downloads are saved to failed_<output_dir>.xlsx for retry
  - Existing files are automatically skipped
        """
    )

    parser.add_argument(
        'excel_file',
        help='Path to Excel file (.xlsx) containing UPC and IMAGES LINK columns'
    )
    parser.add_argument(
        'output_dir',
        help='Output directory for downloaded files'
    )
    parser.add_argument(
        '-t', '--threads',
        type=int,
        default=1,
        metavar='N',
        help='Number of parallel download threads (default: 1)'
    )
    parser.add_argument(
        '-r', '--retry',
        nargs='?',
        const=-1,
        type=int,
        default=0,
        metavar='N',
        help='Auto-retry failed downloads. Use without value for unlimited retries, or specify max attempts'
    )
    parser.add_argument(
        '-d', '--debug',
        action='store_true',
        help='Enable verbose debug output for troubleshooting'
    )
    parser.add_argument(
        '--no-categories',
        action='store_true',
        help='Ignore CATEGORY column and save all files to output directory root'
    )

    return parser.parse_args()


def validate_inputs(args: argparse.Namespace) -> Path:
    """Validate command-line inputs and return the Excel path."""
    excel_path = Path(args.excel_file)
    
    if not excel_path.exists():
        print(f"✗ Error: Excel file not found: {excel_path}")
        sys.exit(1)

    if excel_path.suffix.lower() not in ['.xlsx', '.xls']:
        print("✗ Error: File must be an Excel file (.xlsx or .xls)")
        sys.exit(1)

    if args.threads < 1:
        print("✗ Error: Threads must be at least 1")
        sys.exit(1)

    if args.retry < -1:
        print("✗ Error: Retry value must be -1 (unlimited), 0 (disabled), or a positive number")
        sys.exit(1)

    return excel_path


def run_with_retry(args: argparse.Namespace, excel_path: Path) -> None:
    """Run the batch processor with optional retry logic."""
    # Move import down avoid importing when it is not needed
    from batch_processor import process_excel
    
    current_file = str(excel_path)
    retry_count = 0
    max_retries = args.retry  # -1 = unlimited, 0 = disabled, >0 = limit

    while True:
        failed_excel_path = process_excel(
            excel_file=current_file,
            output_dir=args.output_dir,
            threads=args.threads,
            debug=args.debug,
            no_categories=args.no_categories
        )

        if not failed_excel_path:
            print("\n✅ All downloads completed successfully!")
            break

        # Check if auto-retry is enabled
        if max_retries != 0:
            retry_count += 1

            if max_retries > 0 and retry_count > max_retries:
                print(f"\n⚠️  Reached maximum retry limit ({max_retries} attempts)")
                print(f"📋 Remaining failures saved to: {failed_excel_path.name}")
                _print_retry_hint(failed_excel_path, args)
                break

            if max_retries == -1:
                print(f"\n🔄 Auto-retry #{retry_count} - Retrying failed downloads...\n")
            else:
                print(f"\n🔄 Auto-retry {retry_count}/{max_retries} - Retrying failed downloads...\n")

            current_file = str(failed_excel_path)
            continue

        # Interactive mode
        _handle_interactive_retry(failed_excel_path, args)
        break


def _handle_interactive_retry(failed_excel_path: Path, args: argparse.Namespace) -> None:
    """Handle interactive retry prompt."""
    print("\n" + "=" * 60)
    while True:
        response = input("Would you like to retry the failed downloads now? (Y/N/D): ").strip().upper()
        
        if response in ['Y', 'YES']:
            print("\n🔄 Retrying failed downloads...\n")
            args.retry = -1  # Enable unlimited retry
            run_with_retry(args, failed_excel_path)
            return
        elif response in ['D', 'DEBUG']:
            print("\n🔍 Retrying failed downloads with DEBUG mode enabled...\n")
            args.debug = True
            args.retry = -1
            run_with_retry(args, failed_excel_path)
            return
        elif response in ['N', 'NO']:
            print("\n👋 Exiting. You can retry later by running:")
            _print_retry_hint(failed_excel_path, args)
            return
        else:
            print("   Please enter Y (yes), N (no), or D (debug mode).")


def _print_retry_hint(failed_excel_path: Path, args: argparse.Namespace) -> None:
    """Print hint for retrying failed downloads."""
    print(f"   python cli.py {failed_excel_path.name} {args.output_dir}")
    if args.threads > 1:
        print(f"   python cli.py {failed_excel_path.name} {args.output_dir} --threads {args.threads}")


def main() -> None:
    """Main entry point."""
    args = parse_args()
    excel_path = validate_inputs(args)
    run_with_retry(args, excel_path)


if __name__ == "__main__":
    main()
