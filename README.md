# any2pdf

Convert any file to a PDF - ideal for situations where you are uploading into a tool that only supports PDFs.

Also supports batch conversion of Azure Blob Storage files.

## Supported File Types

| Category | Extensions |
|----------|-----------|
| PDF | `.pdf` (pass-through with optional attachment) |
| Word | `.doc`, `.docx`, `.rtf`, `.odt`, `.txt`, `.dot` |
| Excel | `.xls`, `.xlsx`, `.ods`, `.csv`, `.xlsm` |
| PowerPoint | `.ppt`, `.pptx`, `.odp` |
| Images | `.jpg`, `.jpeg`, `.jfif`, `.png`, `.tif`, `.tiff`, `.bmp`, `.heic` |
| HTML | `.html`, `.htm` |
| Email | `.msg`, `.eml` |

Files with unknown extensions are detected by magic bytes (file signature) when possible.

## Requirements

- Python 3.10+
- Windows (uses COM automation for Office/Outlook)
- Microsoft Office (Word, Excel, PowerPoint, Outlook)
- Microsoft Edge (for HTML rendering)

### Python Dependencies

```bash
pip install -r requirements.txt
```

Key packages:
- `Pillow` + `pillow-heif` - Image conversion including HEIC
- `pypdf` - PDF manipulation and attachments
- `reportlab` - PDF generation
- `pywin32` - COM automation for Office
- `filetype` - Magic byte detection
- `azure-storage-blob`, `azure-identity` - Azure Blob Storage (for migration script)

## Scripts

### `any2pdf.py` - Core Conversion Module

Convert individual files to PDF:

```bash
python any2pdf.py input.docx output_dir/
```

Features:
- Converts Office documents via COM automation
- Renders HTML with Edge headless
- Converts images with Pillow (including HEIC)
- Attaches original file to output PDF
- Magic byte detection for files with wrong/missing extensions

### `migrate_blobs_to_pdf.py` - Azure Blob Migration

Batch convert files from Azure Blob Storage:

```bash
# Analyse what's in the source container
python migrate_blobs_to_pdf.py --analyse

# Check migration progress
python migrate_blobs_to_pdf.py --progress

# Run full migration
python migrate_blobs_to_pdf.py

# Process specific files from a list
python migrate_blobs_to_pdf.py --file-list failed_files.txt --force

# Test with limited files
python migrate_blobs_to_pdf.py --max-files 10
python migrate_blobs_to_pdf.py --filter-extension .msg
python migrate_blobs_to_pdf.py --test-all 5  # 5 of each type
```

Requires `.env` file with Azure configuration:
```
STORAGE_ACCOUNT_NAME=your_account
CONTAINER_NAME=your_container
INPUT_PREFIX=source/folder/
OUTPUT_PREFIX=destination/folder/
```

### `extract_failures.py` - Failure Management

Parse migration logs and manage failure lists:

```bash
# Extract failures from log into categorized files
python extract_failures.py --extract

# Update a failure list by removing successfully processed files
python extract_failures.py --update failed_network_timeout.txt
```

Creates categorized failure lists:
- `failed_network_timeout.txt`
- `failed_auth_expired.txt`
- `failed_msg_com_error.txt`
- `failed_corrupt_image.txt`
- `failed_corrupt_office.txt`
- `failed_password_protected.txt`
- `failed_unsupported_format.txt`

### `download_blobs.py` - Download Files Locally

Download specific blobs for local testing:

```bash
python download_blobs.py file_list.txt output_dir/
```

### `check_folder.py` - List Azure Folders

List top-level folders and file counts in Azure container:

```bash
python check_folder.py
```

## Workflow for Batch Migration

1. **Analyse** source files: `python migrate_blobs_to_pdf.py --analyse`
2. **Run migration**: `python migrate_blobs_to_pdf.py`
3. **Extract failures**: `python extract_failures.py --extract`
4. **Fix issues** and reprocess: `python migrate_blobs_to_pdf.py --file-list failed_X.txt --force`
5. **Update failure lists**: `python extract_failures.py --update failed_X.txt`
6. Repeat until all lists are empty

## Notes

- Unsupported file types create placeholder PDFs with the original file attached
- Placeholder PDFs have metadata: `Subject: FALLBACK`
- Truncated images are handled gracefully
- Zero-byte blobs (folder markers) are skipped
- Logs written to `migration.log`
