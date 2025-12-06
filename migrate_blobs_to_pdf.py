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

# File: info and above (includes successes)
file_handler = logging.FileHandler('migration.log', mode='a')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
logger.addHandler(file_handler)


def save_pdf(pdf_path, target_name, local_dir, container_client, category, overwrite):
    """Save PDF locally or upload to Azure. Returns destination for logging."""
    if local_dir:
        dest_dir = local_dir / category
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / (pathlib.Path(target_name).stem + '.pdf')
        shutil.copy2(pdf_path, dest)
        return dest
    client = container_client.get_blob_client(target_name)
    with open(pdf_path, "rb") as f:
        client.upload_blob(f, overwrite=overwrite)
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
        "--progress",
        action="store_true",
        help="Show migration progress by comparing source and target file counts"
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
    parser.add_argument(
        "--force",
        action="store_true",
        help="Force reprocessing even if output already exists (skips existence check)"
    )
    parser.add_argument(
        "--file-list",
        type=pathlib.Path,
        default=None,
        help="Process only files listed in this text file (one blob path per line, use with --force)"
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

    # Progress mode: compare source and target counts
    if args.progress:
        print("Checking migration progress...")

        # Count source files (excluding directories and zero-byte markers)
        source_files = {}
        source_by_ext = defaultdict(int)
        print(f"  Scanning source: {INPUT_PREFIX}")
        for blob in container_client.list_blobs(name_starts_with=INPUT_PREFIX):
            if blob.name.endswith("/") or blob.size == 0:
                continue
            relative_path = blob.name[len(INPUT_PREFIX):]
            if relative_path.startswith(('Logs/', 'Mapping Tables/')):
                continue
            ext = pathlib.Path(blob.name).suffix.lower()
            source_by_ext[ext] += 1
            # Build expected output name
            if '.' in relative_path.split('/')[-1]:
                expected_output = OUTPUT_PREFIX + relative_path.rsplit('.', 1)[0] + '.pdf'
            else:
                expected_output = OUTPUT_PREFIX + relative_path + '.pdf'
            source_files[expected_output] = blob.name

        # Count target files
        target_files = set()
        print(f"  Scanning target: {OUTPUT_PREFIX}")
        for blob in container_client.list_blobs(name_starts_with=OUTPUT_PREFIX):
            if blob.name.endswith("/") or blob.size == 0:
                continue
            target_files.add(blob.name)

        # Calculate progress
        total_source = len(source_files)
        converted = len(target_files & set(source_files.keys()))
        remaining = total_source - converted
        pct = (converted / total_source * 100) if total_source > 0 else 0

        print(f"\nMigration Progress:")
        print(f"  Source files:    {total_source:,}")
        print(f"  Converted:       {converted:,}")
        print(f"  Remaining:       {remaining:,}")
        print(f"  Progress:        {pct:.1f}%")

        # Show breakdown by extension
        print(f"\nBy extension:")
        for ext in sorted(source_by_ext.keys()):
            count = source_by_ext[ext]
            status = "supported" if ext in ALL_SUPPORTED_EXTENSIONS else "UNSUPPORTED"
            print(f"  {ext or '(none)':>10}: {count:>6} ({status})")

        return

    # List blobs for processing - either from file list or full prefix scan
    if args.file_list:
        # Read specific files from list
        logger.info(f"Loading file list from {args.file_list}")
        with open(args.file_list, 'r', encoding='utf-8') as f:
            file_list = {line.strip() for line in f if line.strip()}
        logger.info(f"Loaded {len(file_list)} files to process")
        
        # Fetch blob properties for each file in the list
        blobs = []
        for blob_name in file_list:
            try:
                blob_client = container_client.get_blob_client(blob_name)
                props = blob_client.get_blob_properties()
                blobs.append(props)
            except Exception as e:
                logger.warning(f"Could not find blob: {blob_name} - {e}")
        
        if not args.force:
            logger.warning("--file-list is typically used with --force to reprocess failed files")
    else:
        # Full prefix scan
        blobs = list(container_client.list_blobs(name_starts_with=INPUT_PREFIX))
    
    # Get existing output files (to skip already-converted)
    existing_outputs = set()
    if not args.force and not OVERWRITE_OUTPUT and not args.local_output:
        logger.info("Loading existing output files...")
        existing_outputs = {b.name for b in container_client.list_blobs(name_starts_with=OUTPUT_PREFIX)}
        logger.info(f"Found {len(existing_outputs)} existing output files")
    elif args.force:
        logger.info("Force mode: skipping existence check, will overwrite")
    
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
            
            # Skip zero-byte files (folder markers in Azure)
            if blob.size == 0:
                continue
            
            # Skip files in excluded directories
            relative_path = blob.name[len(INPUT_PREFIX):]
            if relative_path.startswith(('Logs/', 'Mapping Tables/')):
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
            relative_path = blob.name[len(INPUT_PREFIX):]
            # Change extension to .pdf
            if '.' in relative_path.split('/')[-1]:
                target_pdf_name = OUTPUT_PREFIX + relative_path.rsplit('.', 1)[0] + '.pdf'
            else:
                target_pdf_name = OUTPUT_PREFIX + relative_path + '.pdf'
            
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
                overwrite = args.force or OVERWRITE_OUTPUT
                dest = save_pdf(pdf_path, target_pdf_name, args.local_output, container_client, category, overwrite)
                if not fallback:
                    logger.info(f"OK {category} {category_counts[category]} {blob.name} -> {dest}")
                processed_count += 1
                
            except Exception as e:
                logger.error(f"ERROR {blob.name} : {e}")
                processed_count += 1


if __name__ == "__main__":
    main()
