"""
Module for converting various file formats to PDF.

Supports:
- Office documents (Word, Excel, PowerPoint) via COM automation
- Images via Pillow
- HTML via Microsoft Edge headless
- Outlook messages (.msg) via Outlook COM + Edge headless
- PDF pass-through (no-op)

This module only works with local files and has no Azure/blob storage dependencies.
"""

import argparse
import html
import io
import logging
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile
import traceback
from email import policy
from email.parser import BytesParser
from typing import Optional

import win32com.client
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


# Supported file extensions by category
PDF_EXTENSIONS = {'.pdf'}
WORD_EXTENSIONS = {'.doc', '.docx', '.rtf', '.odt', '.txt', '.dot'}
EXCEL_EXTENSIONS = {'.xls', '.xlsx', '.ods', '.csv', '.xlsm'}
PPT_EXTENSIONS = {'.ppt', '.pptx', '.odp'}
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.heic'}
HTML_EXTENSIONS = {'.html', '.htm'}
MSG_EXTENSIONS = {'.msg'}
EML_EXTENSIONS = {'.eml'}

ALL_SUPPORTED_EXTENSIONS = (
    PDF_EXTENSIONS | WORD_EXTENSIONS | EXCEL_EXTENSIONS | 
    PPT_EXTENSIONS | IMAGE_EXTENSIONS | HTML_EXTENSIONS | MSG_EXTENSIONS |
    EML_EXTENSIONS 
)

# Edge is typically not in PATH, use standard installation location
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# Module-level logger
logger = logging.getLogger(__name__)


def get_category_for_extension(ext: str) -> str:
    """Return the handler category for a file extension."""
    ext = ext.lower()
    if ext in PDF_EXTENSIONS:
        return 'pdf'
    if ext in WORD_EXTENSIONS:
        return 'word'
    if ext in EXCEL_EXTENSIONS:
        return 'excel'
    if ext in PPT_EXTENSIONS:
        return 'ppt'
    if ext in IMAGE_EXTENSIONS:
        return 'image'
    if ext in HTML_EXTENSIONS:
        return 'html'
    if ext in MSG_EXTENSIONS:
        return 'msg'
    if ext in EML_EXTENSIONS:
        return 'eml'
    return 'attachment'


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
        password_args = {'PasswordDocument': password} if password else {}
        
        # Open document
        doc = word.Documents.Open(
            str(src_path.absolute()),
            ReadOnly=True,
            **password_args
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
        password_args = {'Password': password} if password else {}
        
        # Open workbook
        workbook = excel.Workbooks.Open(
            str(src_path.absolute()),
            ReadOnly=True,
            **password_args
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
        password_args = {'Password': password} if password else {}
        
        # Open presentation
        presentation = ppt.Presentations.Open(
            str(src_path.absolute()),
            ReadOnly=True,
            WithWindow=False,
            **password_args
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
    shutil.copy2(src_path, dst_path)
    
    # Note: We don't attach_original for PDFs by default since the original
    # is identical to the output. However, if explicitly requested, we could.
    # The spec says "Do NOT attach_original by default" for PDFs, so we skip it.
    
    return dst_path


def _convert_html_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert HTML to PDF using Microsoft Edge headless."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    # Convert HTML to PDF using Edge headless with temporary user data directory
    temp_user_data = None
    try:
        # Create a temporary user data directory
        temp_user_data = tempfile.mkdtemp(prefix="edge_temp_")
        
        cmd = [
            EDGE_PATH,
            "--headless=new",  # Use new headless mode
            "--disable-gpu",
            "--disable-extensions",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            f"--user-data-dir={temp_user_data}",  # Use temp profile
            f"--print-to-pdf={dst_path.resolve()}",
            str(src_path.resolve()),
        ]
        
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        
        if result.returncode != 0:
            logger.debug(f"Edge STDOUT: {result.stdout.decode(errors='ignore')}")
            logger.debug(f"Edge STDERR: {result.stderr.decode(errors='ignore')}")
            raise RuntimeError(f"Edge headless failed for {src_path} with exit code {result.returncode}")
        
        if not dst_path.exists():
            raise RuntimeError("Edge headless did not create output file")
    
    finally:
        # Clean up temporary user data directory
        if temp_user_data and os.path.exists(temp_user_data):
            shutil.rmtree(temp_user_data, ignore_errors=True)
    
    # Attach original if requested
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def _convert_eml_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert .eml file to PDF via HTML."""
    # Parse the .eml file
    with open(src_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    
    # Extract headers
    from_header = msg.get('From', '')
    to_header = msg.get('To', '')
    subject_header = msg.get('Subject', '')
    date_header = msg.get('Date', '')
    
    # Find body content - prefer HTML, fall back to plain text
    body_content = None
    is_html = False
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html' and body_content is None:
                body_content = part.get_content()
                is_html = True
                break
        if body_content is None:
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain':
                    body_content = part.get_content()
                    is_html = False
                    break
    else:
        content_type = msg.get_content_type()
        if content_type == 'text/html':
            body_content = msg.get_content()
            is_html = True
        elif content_type == 'text/plain':
            body_content = msg.get_content()
            is_html = False
    
    if body_content is None:
        raise ValueError(f"No text/html or text/plain body found in {src_path}")
    
    # Build HTML document
    header_html = f"""<div style="font-family: Arial, sans-serif; border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 20px;">
<p><strong>From:</strong> {html.escape(from_header)}</p>
<p><strong>To:</strong> {html.escape(to_header)}</p>
<p><strong>Subject:</strong> {html.escape(subject_header)}</p>
<p><strong>Date:</strong> {html.escape(date_header)}</p>
</div>"""
    
    if is_html:
        body_html = body_content
    else:
        body_html = f"<pre>{html.escape(body_content)}</pre>"
    
    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{html.escape(subject_header)}</title>
</head>
<body>
{header_html}
{body_html}
</body>
</html>"""
    
    # Write temporary HTML file
    temp_html = dst_dir / f"{src_path.stem}_temp.html"
    temp_html.write_text(full_html, encoding='utf-8')
    
    try:
        # Convert HTML to PDF
        pdf_path = _convert_html_to_pdf(temp_html, dst_dir, attach_original=False)
        
        # Rename to match original .eml filename
        final_path = dst_dir / f"{src_path.stem}.pdf"
        if pdf_path != final_path:
            os.replace(pdf_path, final_path)
        
        # Attach original if requested
        if attach_original:
            attach_original_to_pdf(final_path, src_path)
        
        return final_path
    finally:
        # Clean up temp HTML file
        try:
            temp_html.unlink()
        except OSError:
            pass


def _convert_msg_to_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Convert Outlook .msg file to PDF via HTML."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    outlook = None
    msg = None
    temp_html = None
    
    try:
        # Create Outlook application instance
        outlook = win32com.client.DispatchEx("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Open the .msg file
        msg = namespace.OpenSharedItem(str(src_path.absolute()))
        
        # Save as HTML to temporary file
        temp_html = dst_dir / f"{src_path.stem}_temp.html"
        msg.SaveAs(str(temp_html.absolute()), 5)  # olHTML = 5
        
        # Convert the HTML to PDF using Edge (creates PDF with temp HTML name)
        temp_pdf = _convert_html_to_pdf(temp_html, dst_dir, attach_original=False)
        
        # Rename the PDF to match the original .msg filename
        if temp_pdf != dst_path:
            os.replace(temp_pdf, dst_path)
        
    finally:
        # Clean up
        if msg:
            msg.Close(0)  # olDiscard = 0
        if outlook:
            outlook.Quit()
        
        # Remove temporary HTML file
        if temp_html and temp_html.exists():
            temp_html.unlink()
    
    # Attach original .msg file
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
    return dst_path


def create_placeholder_pdf(
    src_path: pathlib.Path,
    dst_dir: pathlib.Path,
    attach_original: bool
) -> pathlib.Path:
    """Create a placeholder PDF with the original file attached (for non-convertible files)."""
    dst_path = dst_dir / f"{src_path.stem}.pdf"
    
    # Create a simple PDF with a placeholder message
    # Create PDF content
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    
    # Add text to the page
    c.setFont("Helvetica", 12)
    c.drawString(100, 750, f"Original file: {src_path.name}")
    c.drawString(100, 730, f"File type: {src_path.suffix}")
    c.drawString(100, 700, "This file type cannot be converted to PDF.")
    c.drawString(100, 680, "The original file is attached to this PDF.")
    
    c.save()
    
    # Write the PDF
    buffer.seek(0)
    with open(dst_path, 'wb') as f:
        f.write(buffer.read())
    
    # Attach original file
    if attach_original:
        attach_original_to_pdf(dst_path, src_path)
    
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
    
    elif ext in HTML_EXTENSIONS:
        return _convert_html_to_pdf(src_path, dst_dir, attach_original)
    
    elif ext in MSG_EXTENSIONS:
        return _convert_msg_to_pdf(src_path, dst_dir, attach_original)
    
    elif ext in EML_EXTENSIONS:
        return _convert_eml_to_pdf(src_path, dst_dir, attach_original)
    
    else:
        # Unsupported file extension - raise error
        raise ValueError(
            f"Unsupported file extension: '{ext}'. "
            f"Supported extensions are: {', '.join(sorted(ALL_SUPPORTED_EXTENSIONS))}"
        )


def main():
    """Main CLI entry point."""
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(levelname)s: %(message)s'
    )
    
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
        logger.error(f"Input file does not exist: {args.input}")
        return 1
    
    # Check if extension is supported
    if args.input.suffix.lower() not in ALL_SUPPORTED_EXTENSIONS:
        logger.error(f"Unsupported file extension: {args.input.suffix}")
        logger.error(f"Supported: {', '.join(sorted(ALL_SUPPORTED_EXTENSIONS))}")
        return 1
    
    if args.verbose:
        logger.setLevel(logging.DEBUG)
        logger.info(f"Input: {args.input}")
        logger.info(f"Output directory: {args.output_dir}")
        logger.info(f"Attach original: {not args.no_attach_original}")
        logger.info(f"File size: {args.input.stat().st_size:,} bytes")
    
    try:
        result = convert_anything_to_pdf(
            src_path=args.input,
            dst_dir=args.output_dir,
            attach_original=not args.no_attach_original
        )

        logger.info("Success!")

        if args.verbose:
            logger.info(f"PDF created: {result}")
            logger.info(f"PDF size: {result.stat().st_size:,} bytes")
        
        return 0
        
    except Exception as e:
        logger.error(f"{e}")
        if args.verbose:
            logger.exception("Full traceback:")
        return 1


if __name__ == '__main__':
    sys.exit(main())
