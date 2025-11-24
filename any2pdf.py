"""
Module for converting various file formats to PDF.

Supports:
- Office documents (Word, Excel, PowerPoint) via COM automation
- Images via Pillow
- PDF pass-through (no-op)

This module only works with local files and has no Azure/blob storage dependencies.
"""

import os
import pathlib
import tempfile
from typing import Optional

import win32com.client
from PIL import Image
from pypdf import PdfReader, PdfWriter


# Supported file extensions by category
PDF_EXTENSIONS = {'.pdf'}
WORD_EXTENSIONS = {'.doc', '.docx', '.rtf', '.odt'}
EXCEL_EXTENSIONS = {'.xls', '.xlsx', '.ods'}
PPT_EXTENSIONS = {'.ppt', '.pptx', '.odp'}
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp'}

ALL_SUPPORTED_EXTENSIONS = (
    PDF_EXTENSIONS | WORD_EXTENSIONS | EXCEL_EXTENSIONS | 
    PPT_EXTENSIONS | IMAGE_EXTENSIONS
)


def get_password_for_file(path: pathlib.Path) -> Optional[str]:
    """
    Retrieve password for a protected document.
    
    Currently reads from MIGRATION_DOC_PASSWORD environment variable.
    Returns None if not set.
    
    Args:
        path: Path to the file (reserved for future per-file password logic)
        
    Returns:
        Password string or None
    """
    return os.environ.get('MIGRATION_DOC_PASSWORD')


def attach_original_to_pdf(pdf_path: pathlib.Path, original_path: pathlib.Path) -> None:
    """
    Attach the original file as an embedded attachment in the PDF.
    
    Args:
        pdf_path: Path to the PDF file to modify
        original_path: Path to the original file to embed
        
    Raises:
        IOError: If PDF cannot be read or written
    """
    # Read the existing PDF
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    
    # Copy all pages
    for page in reader.pages:
        writer.add_page(page)
    
    # Read the original file as bytes
    with open(original_path, 'rb') as f:
        original_bytes = f.read()
    
    # Attach the original file
    writer.add_attachment(original_path.name, original_bytes)
    
    # Write to a temporary file, then atomically replace
    with tempfile.NamedTemporaryFile(
        mode='wb',
        delete=False,
        dir=pdf_path.parent,
        suffix='.pdf'
    ) as tmp:
        tmp_path = pathlib.Path(tmp.name)
        writer.write(tmp)
    
    # Atomic replace
    os.replace(tmp_path, pdf_path)


def _convert_word_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert Word document to PDF using COM automation."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    word = None
    doc = None
    try:
        # Create Word application instance
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        # Get password if available
        password = get_password_for_file(src_path)
        
        # Open document
        try:
            if password:
                doc = word.Documents.Open(
                    str(src_path.absolute()),
                    ReadOnly=True,
                    PasswordDocument=password
                )
            else:
                doc = word.Documents.Open(
                    str(src_path.absolute()),
                    ReadOnly=True
                )
        except Exception as e:
            raise RuntimeError(
                f"Failed to open Word document '{src_path}'. "
                f"It may be password-protected or corrupted: {e}"
            )
        
        # Export to PDF (wdExportFormatPDF = 17)
        doc.ExportAsFixedFormat(
            str(dst_path.absolute()),
            17,  # wdExportFormatPDF
            OpenAfterExport=False
        )
        
    finally:
        # Clean up
        if doc:
            doc.Close(False)
        if word:
            word.Quit()
    
    # Attach original if requested
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def _convert_excel_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert Excel workbook to PDF using COM automation."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    excel = None
    workbook = None
    try:
        # Create Excel application instance
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Get password if available
        password = get_password_for_file(src_path)
        
        # Open workbook
        try:
            if password:
                workbook = excel.Workbooks.Open(
                    str(src_path.absolute()),
                    ReadOnly=True,
                    Password=password
                )
            else:
                workbook = excel.Workbooks.Open(
                    str(src_path.absolute()),
                    ReadOnly=True
                )
        except Exception as e:
            raise RuntimeError(
                f"Failed to open Excel workbook '{src_path}'. "
                f"It may be password-protected or corrupted: {e}"
            )
        
        # Export to PDF (xlTypePDF = 0)
        workbook.ExportAsFixedFormat(
            0,  # xlTypePDF
            str(dst_path.absolute())
        )
        
    finally:
        # Clean up
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()
    
    # Attach original if requested
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def _convert_ppt_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert PowerPoint presentation to PDF using COM automation."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    ppt = None
    presentation = None
    try:
        # Create PowerPoint application instance
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        # PowerPoint doesn't have Visible property in the same way
        # but we can keep it minimal by not showing windows
        
        # Get password if available
        password = get_password_for_file(src_path)
        
        # Open presentation
        try:
            if password:
                presentation = ppt.Presentations.Open(
                    str(src_path.absolute()),
                    ReadOnly=True,
                    WithWindow=False,
                    OpenAndRepair=False,
                    Password=password
                )
            else:
                presentation = ppt.Presentations.Open(
                    str(src_path.absolute()),
                    ReadOnly=True,
                    WithWindow=False
                )
        except Exception as e:
            raise RuntimeError(
                f"Failed to open PowerPoint presentation '{src_path}'. "
                f"It may be password-protected or corrupted: {e}"
            )
        
        # Save as PDF (ppSaveAsPDF = 32)
        presentation.SaveAs(
            str(dst_path.absolute()),
            32  # ppSaveAsPDF
        )
        
    finally:
        # Clean up
        if presentation:
            presentation.Close()
        if ppt:
            ppt.Quit()
    
    # Attach original if requested
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def _convert_image_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert image to PDF using Pillow."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    # Open image
    img = Image.open(src_path)
    
    # Convert to RGB if necessary
    if img.mode in ('RGBA', 'P'):
        # Create a white background
        if img.mode == 'P':
            img = img.convert('RGBA')
        
        background = Image.new('RGB', img.size, (255, 255, 255))
        if img.mode == 'RGBA':
            background.paste(img, mask=img.split()[3])  # Use alpha channel as mask
        else:
            background.paste(img)
        img = background
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Save as PDF
    img.save(dst_path, 'PDF', resolution=100.0)
    
    # Attach original if requested
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def _handle_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Handle PDF files (no-op or copy)."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    # If source and destination are the same, just return it
    if src_path.resolve() == dst_path.resolve():
        return src_path
    
    # Copy the PDF to the destination
    import shutil
    shutil.copy2(src_path, dst_path)
    
    # Note: We don't attach_original for PDFs by default since the original
    # is identical to the output. However, if explicitly requested, we could.
    # The spec says "Do NOT attach_original by default" for PDFs, so we skip it.
    
    return dst_path


def convert_anything_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool = True,
) -> pathlib.Path:
    """
    Convert a local file to PDF.
    
    Supports Office documents (via COM), images (via Pillow), and PDF pass-through.
    
    Args:
        src_path: Path to the source file
        dst_dir: Directory where the output PDF will be created
        attach_original: If True, embed the original file into the PDF
                        (except for PDF inputs, where it's skipped by default)
    
    Returns:
        Path to the generated PDF file
        
    Raises:
        ValueError: If the file extension is not supported
        RuntimeError: If conversion fails (e.g., password-protected file without password)
        IOError: If file cannot be read or written
    """
    # Ensure source file exists
    if not src_path.exists():
        raise FileNotFoundError(f"Source file not found: {src_path}")
    
    # Create destination directory if it doesn't exist
    dst_dir.mkdir(parents=True, exist_ok=True)
    
    # Get file extension (lowercase)
    ext = src_path.suffix.lower()
    
    # Route to appropriate converter
    if ext in PDF_EXTENSIONS:
        # For PDFs, don't attach original by default
        return _handle_pdf(src_path, dst_dir, attach_original=False)
    
    elif ext in WORD_EXTENSIONS:
        return _convert_word_to_pdf(src_path, dst_dir, attach_original)
    
    elif ext in EXCEL_EXTENSIONS:
        return _convert_excel_to_pdf(src_path, dst_dir, attach_original)
    
    elif ext in PPT_EXTENSIONS:
        return _convert_ppt_to_pdf(src_path, dst_dir, attach_original)
    
    elif ext in IMAGE_EXTENSIONS:
        return _convert_image_to_pdf(src_path, dst_dir, attach_original)
    
    else:
        raise ValueError(
            f"Unsupported file extension: '{ext}'. "
            f"Supported extensions are: {', '.join(sorted(ALL_SUPPORTED_EXTENSIONS))}"
        )


def main():
    """Main CLI entry point."""
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(
        description="Convert various file formats to PDF",
        epilog=f"Supported extensions: {', '.join(sorted(ALL_SUPPORTED_EXTENSIONS))}"
    )
    
    parser.add_argument(
        'input',
        type=pathlib.Path,
        help='Path to the input file to convert'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        type=pathlib.Path,
        default=pathlib.Path('output'),
        help='Directory for output PDF (default: ./output)'
    )
    
    parser.add_argument(
        '--no-attach-original',
        action='store_true',
        help='Do not embed the original file in the PDF'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Print verbose output'
    )
    
    args = parser.parse_args()
    
    # Validate input file exists
    if not args.input.exists():
        print(f"Error: Input file does not exist: {args.input}", file=sys.stderr)
        return 1
    
    # Check if extension is supported
    if args.input.suffix.lower() not in ALL_SUPPORTED_EXTENSIONS:
        print(f"Error: Unsupported file extension: {args.input.suffix}", file=sys.stderr)
        print(f"Supported: {', '.join(sorted(ALL_SUPPORTED_EXTENSIONS))}", file=sys.stderr)
        return 1
    
    if args.verbose:
        print(f"Input: {args.input}")
        print(f"Output directory: {args.output_dir}")
        print(f"Attach original: {not args.no_attach_original}")
        print(f"File size: {args.input.stat().st_size:,} bytes")
    
    try:
        result = convert_anything_to_pdf(
            src_path=args.input,
            dst_dir=args.output_dir,
            attach_original=not args.no_attach_original
        )
        
        if args.verbose:
            print(f"\nSuccess!")
            print(f"PDF created: {result}")
            print(f"PDF size: {result.stat().st_size:,} bytes")
        else:
            print(result)
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == '__main__':
    import sys
    sys.exit(main())
