# Dropbox Downloader

Batch download "the first file" from Dropbox shared folders using Excel file input.

## Usage

Install dependencies first (you should use venv):
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Create an Excel file (`.xlsx`) with these columns (column's order does not matter):

- `UPC` - Unique identifier (used as filename)
- `IMAGES LINK` - Dropbox shared folder URL
- `CATEGORY` (optional) - Organize files into subdirectories

See `products.xlsx` for an example Excel file.

<img width="2670" height="380" alt="Excel file format" src="https://github.com/user-attachments/assets/c9ad03de-63e0-4637-9e9b-4eb719b1bf33" />

The **Dropbox folder link** should look like this when accessed:

<img width="568" height="626" alt="Dropbox download folder" src="https://github.com/user-attachments/assets/09f0d07a-1c51-4493-934b-87c2236dd4d4" />


### 3. Run Batch Download

```bash
# Single-threaded download
python cli.py products.xlsx output/

# Multi-threaded download (4 threads)
python cli.py products.xlsx output/ --threads 4

# Auto-retry failed downloads
python cli.py products.xlsx output/ --retry

# Enable debug mode
python cli.py products.xlsx output/ --debug
```

### Command-Line Flags

| Flag | Description | Example |
|------|-------------|---------|
| `-t, --threads N` | Number of parallel download threads | `--threads 4` |
| `-r, --retry [N]` | Auto-retry failed downloads (unlimited if no value, or max N attempts) | `--retry` or `--retry 3` |
| `-d, --debug` | Enable verbose debug output | `--debug` |
| `--no-categories` | Ignore CATEGORY column and save all files to root output directory | `--no-categories` |
| `-h, --help` | Show help message | `--help` |

## Features

- **Multi-threaded downloads** - Download multiple files in parallel
- **Auto-retry** - Automatically retry failed downloads
- **Skip existing files** - Avoids re-downloading files that already exist
- **Progress tracking** - Real-time progress bars with tqdm
- **Failed download tracking** - Creates `failed_*.xlsx` for easy retry

## Requirements

- Python 3.10+ (I tested this on Python 3.10)
- Chrome browser (for Selenium)
- ChromeDriver (should be installed automatically with Selenium)

## Tech Stack

Python, Pandas, Selenium, ChromeDriver, tqdm
