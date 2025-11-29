"""
One-off migration script to convert Azure Blob Storage files to PDF.

This script downloads blobs from a specified prefix, converts them to PDF
using any2pdf.convert_anything_to_pdf, and uploads the results to a new prefix.

Authentication uses Azure AD (DefaultAzureCredential). Run `az login` first.

Configuration is loaded from .env file.
"""

import argparse
import logging
import os
import pathlib
import shutil
import tempfile
from collections import defaultdict
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from any2pdf import (
    convert_anything_to_pdf, ALL_SUPPORTED_EXTENSIONS, create_placeholder_pdf,
    get_category_for_extension
)


# ============================================================================
# CONFIGURATION (Load from .env)
# ============================================================================

load_dotenv()

STORAGE_ACCOUNT_NAME = os.getenv("STORAGE_ACCOUNT_NAME")
CONTAINER_NAME = os.getenv("CONTAINER_NAME")
INPUT_PREFIX = os.getenv("INPUT_PREFIX")
OUTPUT_PREFIX = os.getenv("OUTPUT_PREFIX")
OVERWRITE_OUTPUT = os.getenv("OVERWRITE_OUTPUT", "False").lower() in ("true", "1", "yes")

# Logging setup
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Console: show all
console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
console.setFormatter(logging.Formatter('%(message)s'))
logger.addHandler(console)

# File: warnings and errors only
file_handler = logging.FileHandler('migration_issues.log', mode='w')
file_handler.setLevel(logging.WARNING)
file_handler.setFormatter(logging.Formatter('%(levelname)s %(message)s'))
logger.addHandler(file_handler)


def save_pdf(pdf_path, target_name, local_dir, container_client, category):
    """Save PDF locally or upload to Azure. Returns destination for logging."""
    if local_dir:
        dest_dir = local_dir / category
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / (pathlib.Path(target_name).stem + '.pdf')
        shutil.copy2(pdf_path, dest)
        return dest
    client = container_client.get_blob_client(target_name)
    with open(pdf_path, "rb") as f:
        client.upload_blob(f, overwrite=OVERWRITE_OUTPUT)
    return target_name


# ============================================================================
# Main Logic
# ============================================================================

def main():
    """Download blobs, convert to PDF, upload results."""
    
    # Validate configuration
    if not all([STORAGE_ACCOUNT_NAME, CONTAINER_NAME, INPUT_PREFIX, OUTPUT_PREFIX]):
        logger.error("Missing required configuration in .env file")
        logger.error("Required: STORAGE_ACCOUNT_NAME, CONTAINER_NAME, INPUT_PREFIX, OUTPUT_PREFIX")
        return
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description="Migrate Azure blobs to PDF format"
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=None,
        help="Maximum number of files to process (useful for testing, default: process all)"
    )
    parser.add_argument(
        "--analyse",
        action="store_true",
        help="Analyse file extensions and show what's supported (doesn't process files)"
    )
    parser.add_argument(
        "--filter-extension",
        type=str,
        default=None,
        help="Only process files with this extension (e.g. '.msg' or '.docx')"
    )
    parser.add_argument(
        "--test-all",
        type=int,
        metavar="N",
        default=None,
        help="Test mode: process N files of each supported extension type"
    )
    parser.add_argument(
        "--local-output",
        type=pathlib.Path,
        default=None,
        help="Write PDFs to local directory instead of uploading to Azure (for testing)"
    )
    args = parser.parse_args()
    
    # Authenticate with Azure AD
    credential = DefaultAzureCredential()
    blob_service = BlobServiceClient(
        f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net",
        credential=credential,
    )
    
    container_client = blob_service.get_container_client(CONTAINER_NAME)
    
    # If analysis mode, just report extensions and exit
    if args.analyse:
        print("Analysing file extensions in source...")
        extension_stats = defaultdict(int)
        unsupported_files = []
        total_files = 0
        
        blobs = container_client.list_blobs(name_starts_with=INPUT_PREFIX)
        
        for blob in blobs:
            if blob.name.endswith("/"):
                continue
            
            total_files += 1
            ext = pathlib.Path(blob.name).suffix.lower()
            extension_stats[ext] += 1
            
            if ext not in ALL_SUPPORTED_EXTENSIONS:
                unsupported_files.append(blob.name)
        
        # Report findings
        supported_count = total_files - len(unsupported_files)
        print(f"Found {total_files} files ({supported_count} supported, {len(unsupported_files)} unsupported):")
        
        # Show supported extensions
        print("Supported extensions:")
        for ext, count in sorted(extension_stats.items()):
            if ext in ALL_SUPPORTED_EXTENSIONS:
                print(f"  {ext or '(no extension)'}: {count} file(s)")
        
        # Show unsupported extensions
        unsupported_extensions = {ext: count for ext, count in extension_stats.items() if ext not in ALL_SUPPORTED_EXTENSIONS}
        if unsupported_extensions:
            print("Unsupported extensions:")
            for ext, count in sorted(unsupported_extensions.items()):
                print(f"  {ext or '(no extension)'}: {count} file(s)")
        
        if unsupported_files:
            print(f"WARNING: {len(unsupported_files)} unsupported file(s) will be skipped:")
            for fname in unsupported_files[:10]:  # Show first 10
                print(f"  - {fname}")
            if len(unsupported_files) > 10:
                print(f"  ... and {len(unsupported_files) - 10} more")
        
        return
    
    # List all blobs for processing
    blobs = list(container_client.list_blobs(name_starts_with=INPUT_PREFIX))
    
    # Get existing output files (to skip already-converted)
    existing_outputs = set()
    if not OVERWRITE_OUTPUT and not args.local_output:
        logger.info("Loading existing output files...")
        existing_outputs = {b.name for b in container_client.list_blobs(name_starts_with=OUTPUT_PREFIX)}
        logger.info(f"Found {len(existing_outputs)} existing output files")
    
    # Track counts per category
    category_counts = defaultdict(int)
    
    # Create temporary directory for downloads and conversions
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = pathlib.Path(tmpdir)
        
        processed_count = 0
        
        for blob in blobs:
            # Skip directory markers
            if blob.name.endswith("/"):
                continue
            
            # Check max files limit
            if args.max_files is not None and processed_count >= args.max_files:
                logger.info(f"Reached max files limit ({args.max_files}), stopping")
                break
            
            # Determine local and target names
            local_name = pathlib.Path(blob.name).name
            ext = pathlib.Path(blob.name).suffix.lower()
            
            # Filter by extension if specified (do this first, silently)
            if args.filter_extension and ext != args.filter_extension.lower():
                continue
            
            # Get category for this file type
            category = get_category_for_extension(ext)
            
            # Test-all mode: skip if we've already processed N of this category
            if args.test_all and category_counts[category] >= args.test_all:
                continue
            
            category_counts[category] += 1
            
            local_path = tmpdir_path / local_name
            
            # Preserve directory structure: replace INPUT_PREFIX with OUTPUT_PREFIX
            relative_path = pathlib.Path(blob.name[len(INPUT_PREFIX):])
            target_pdf_name = OUTPUT_PREFIX + str(relative_path.with_suffix('.pdf'))
            
            # Skip if already converted
            if target_pdf_name in existing_outputs:
                logger.debug(f"SKIP {category} {category_counts[category]} {target_pdf_name} (already exists)")
                continue
            
            try:
                # Download blob
                downloader = container_client.download_blob(blob.name)
                with open(local_path, "wb") as f:
                    f.write(downloader.readall())
                
                # Convert to PDF (fallback to placeholder on any error)
                fallback = False
                try:
                    pdf_path = convert_anything_to_pdf(
                        local_path,
                        dst_dir=tmpdir_path,
                        attach_original=True
                    )
                except Exception as e:
                    logger.warning(f"FALLBACK {blob.name} : {e}")
                    pdf_path = create_placeholder_pdf(local_path, tmpdir_path, attach_original=True)
                    fallback = True
                
                # Save PDF
                dest = save_pdf(pdf_path, target_pdf_name, args.local_output, container_client, category)
                if not fallback:
                    logger.info(f"OK {category} {category_counts[category]} {blob.name} -> {dest}")
                processed_count += 1
                
            except Exception as e:
                logger.error(f"ERROR {blob.name} : {e}")
                processed_count += 1


if __name__ == "__main__":
    main()
