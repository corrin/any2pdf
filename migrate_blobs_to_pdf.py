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
import shutil
import tempfile
from collections import defaultdict
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from any2pdf import convert_anything_to_pdf, ALL_SUPPORTED_EXTENSIONS, create_placeholder_pdf


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
    parser.add_argument(
        "--filter-extension",
        type=str,
        default=None,
        help="Only process files with this extension (e.g. '.msg' or '.docx')"
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
            
            # Filter by extension if specified (do this first, silently)
            if args.filter_extension and ext != args.filter_extension.lower():
                continue
            
            local_path = tmpdir_path / local_name
            
            # Preserve directory structure: replace INPUT_PREFIX with OUTPUT_PREFIX
            relative_path = blob.name[len(INPUT_PREFIX):]  # Remove INPUT_PREFIX
            relative_path_obj = pathlib.Path(relative_path)
            target_pdf_name = OUTPUT_PREFIX + str(relative_path_obj.with_suffix('.pdf'))
            
            try:
                if args.local_output:
                    # Local output mode - write to local directory
                    local_output_dir = args.local_output / pathlib.Path(relative_path).parent
                    local_output_dir.mkdir(parents=True, exist_ok=True)
                    final_pdf_path = args.local_output / relative_path_obj.with_suffix('.pdf')
                else:
                    # Azure output mode - check if already exists
                    out_client = container_client.get_blob_client(target_pdf_name)
                    
                    if not OVERWRITE_OUTPUT:
                        try:
                            out_client.get_blob_properties()
                            print("SKIP", target_pdf_name, "(already exists)")
                            continue
                        except Exception:
                            pass  # Blob doesn't exist, proceed
                
                # Download blob
                downloader = container_client.download_blob(blob.name)
                with open(local_path, "wb") as f:
                    f.write(downloader.readall())
                
                # Convert to PDF (fallback to placeholder on any error)
                try:
                    pdf_path = convert_anything_to_pdf(
                        local_path,
                        dst_dir=tmpdir_path,
                        attach_original=True
                    )
                except Exception as e:
                    print("FALLBACK", blob.name, ":", e)
                    pdf_path = create_placeholder_pdf(local_path, tmpdir_path, attach_original=True)
                
                # Upload PDF or save locally
                if args.local_output:
                    shutil.copy2(pdf_path, final_pdf_path)
                    print("OK", blob.name, "->", final_pdf_path)
                else:
                    with open(pdf_path, "rb") as f:
                        out_client.upload_blob(f, overwrite=OVERWRITE_OUTPUT)
                    print("OK", blob.name, "->", target_pdf_name)
                
                processed_count += 1
                
            except Exception as e:
                print("ERROR", blob.name, ":", e)
                processed_count += 1


if __name__ == "__main__":
    main()
