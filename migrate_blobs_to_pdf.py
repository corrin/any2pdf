"""
One-off migration script to convert Azure Blob Storage files to PDF.

This script downloads blobs from a specified prefix, converts them to PDF
using any2pdf.convert_anything_to_pdf, and uploads the results to a new prefix.

Authentication uses Azure AD (DefaultAzureCredential). Run `az login` first.

Configuration is loaded from .env file.
"""

import argparse
import os
import pathlib
import tempfile
from collections import defaultdict
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from any2pdf import convert_anything_to_pdf, ALL_SUPPORTED_EXTENSIONS


# ============================================================================
# CONFIGURATION (Load from .env)
# ============================================================================

load_dotenv()

STORAGE_ACCOUNT_NAME = os.getenv("STORAGE_ACCOUNT_NAME")
CONTAINER_NAME = os.getenv("CONTAINER_NAME")
INPUT_PREFIX = os.getenv("INPUT_PREFIX")
OUTPUT_PREFIX = os.getenv("OUTPUT_PREFIX")
OVERWRITE_OUTPUT = os.getenv("OVERWRITE_OUTPUT", "False").lower() in ("true", "1", "yes")


# ============================================================================
# Main Logic
# ============================================================================

def main():
    """Download blobs, convert to PDF, upload results."""
    
    # Validate configuration
    if not all([STORAGE_ACCOUNT_NAME, CONTAINER_NAME, INPUT_PREFIX, OUTPUT_PREFIX]):
        print("ERROR: Missing required configuration in .env file")
        print("Required: STORAGE_ACCOUNT_NAME, CONTAINER_NAME, INPUT_PREFIX, OUTPUT_PREFIX")
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
                print(f"Reached max files limit ({args.max_files}), stopping")
                break
            
            # Determine local and target names
            local_name = pathlib.Path(blob.name).name
            ext = pathlib.Path(blob.name).suffix.lower()
            
            # Skip unsupported file types
            if ext not in ALL_SUPPORTED_EXTENSIONS:
                print("SKIP", blob.name, "(unsupported extension)")
                continue
            
            local_path = tmpdir_path / local_name
            
            # Preserve directory structure: replace INPUT_PREFIX with OUTPUT_PREFIX
            relative_path = blob.name[len(INPUT_PREFIX):]  # Remove INPUT_PREFIX
            relative_path_obj = pathlib.Path(relative_path)
            target_pdf_name = OUTPUT_PREFIX + str(relative_path_obj.with_suffix('.pdf'))
            
            # Check if output already exists (skip if not overwriting)
            if not OVERWRITE_OUTPUT:
                out_client = container_client.get_blob_client(target_pdf_name)
                if out_client.exists():
                    print("SKIP", target_pdf_name)
                    continue
            
            try:
                # Download blob
                downloader = container_client.download_blob(blob.name)
                with open(local_path, "wb") as f:
                    f.write(downloader.readall())
                
                # Convert to PDF
                pdf_path = convert_anything_to_pdf(
                    local_path,
                    dst_dir=tmpdir_path,
                    attach_original=True
                )
                
                # Upload PDF
                out_client = container_client.get_blob_client(target_pdf_name)
                with open(pdf_path, "rb") as f:
                    out_client.upload_blob(f, overwrite=OVERWRITE_OUTPUT)
                
                print("OK", blob.name, "->", target_pdf_name)
                processed_count += 1
                
            except Exception as e:
                print("ERROR", blob.name, ":", e)
                processed_count += 1
                continue


if __name__ == "__main__":
    main()
