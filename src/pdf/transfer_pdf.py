"""
Asset Transfer Form (ATF) PDF Generator
Generates PDF forms matching the Ajman University Main Store format
With Adobe Acrobat Digital Signature field support (Certificate-based)
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
import os
import io

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DictionaryObject,
    ArrayObject,
    NameObject,
    NumberObject,
    TextStringObject,
)

from ..utils.signature import get_form_date, get_logo_path, get_output_path


# Colors matching the original form
BLUE = colors.HexColor("#0098DA")
BLACK = colors.black
WHITE = colors.white


class TransferPDFGenerator:
    """Generates Asset Transfer Form (ATF) PDFs with digital signature field"""

    def __init__(self):
        self.width, self.height = A4
        self.margin = 0.75 * inch
        self.sig_field_rect = None  # Will store signature field coordinates

    def _generate_filename(self, form_data: dict) -> str:
        """Generate filename: Asset Transfer - From {emp id}-{name} to {emp id}-{name}.pdf"""
        from_emp_id = form_data.get("from_emp_id", "").strip() or "Unknown"
        from_name = form_data.get("from_name", "").strip() or "Unknown"
        to_emp_id = form_data.get("to_emp_id", "").strip() or "Unknown"
        to_name = form_data.get("to_name", "").strip() or "Unknown"

        # Clean characters that aren't allowed in filenames
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            from_emp_id = from_emp_id.replace(char, '_')
            from_name = from_name.replace(char, '_')
            to_emp_id = to_emp_id.replace(char, '_')
            to_name = to_name.replace(char, '_')

        return f"Asset Transfer - From {from_emp_id}-{from_name} to {to_emp_id}-{to_name}.pdf"

    def _get_unique_filepath(self, output_dir: str, filename: str) -> str:
        """Get unique filepath, adding (#2), (#3), etc. if file exists"""
        filepath = os.path.join(output_dir, filename)

        if not os.path.exists(filepath):
            return filepath

        # File exists, find unique name
        base_name = filename[:-4]  # Remove .pdf
        counter = 2

        while True:
            new_filename = f"{base_name} (#{counter}).pdf"
            new_filepath = os.path.join(output_dir, new_filename)
            if not os.path.exists(new_filepath):
                return new_filepath
            counter += 1

    def generate(self, form_data: dict, filename: str = None) -> str:
        """Generate the PDF form with digital signature field"""
        if filename is None:
            filename = self._generate_filename(form_data)

        output_dir = get_output_path()
        os.makedirs(output_dir, exist_ok=True)
        filepath = self._get_unique_filepath(output_dir, filename)

        # First, create the base PDF in memory
        pdf_buffer = io.BytesIO()
        c = canvas.Canvas(pdf_buffer, pagesize=A4)
        self._draw_form(c, form_data)
        c.save()

        # Now add the signature field using pypdf
        pdf_buffer.seek(0)
        self._add_signature_field(pdf_buffer, filepath)

        return filepath

    def _add_signature_field(self, pdf_buffer: io.BytesIO, output_path: str):
        """Add a proper digital signature field to the PDF"""
        reader = PdfReader(pdf_buffer)
        writer = PdfWriter()

        # Copy all pages
        for page in reader.pages:
            writer.add_page(page)

        # Get the last page (where signature is)
        last_page = writer.pages[-1]

        # Create signature field annotation
        if self.sig_field_rect:
            x1, y1, x2, y2 = self.sig_field_rect

            # Create the signature field dictionary
            sig_field = DictionaryObject()
            sig_field.update({
                NameObject("/Type"): NameObject("/Annot"),
                NameObject("/Subtype"): NameObject("/Widget"),
                NameObject("/FT"): NameObject("/Sig"),
                NameObject("/T"): TextStringObject("Signature"),
                NameObject("/F"): NumberObject(4),  # Print flag
                NameObject("/Rect"): ArrayObject([
                    NumberObject(int(x1)),
                    NumberObject(int(y1)),
                    NumberObject(int(x2)),
                    NumberObject(int(y2))
                ]),
                NameObject("/P"): last_page.indirect_reference,
            })

            # Add annotation to page
            if "/Annots" not in last_page:
                last_page[NameObject("/Annots")] = ArrayObject()

            # Add the signature field to annotations
            sig_field_ref = writer._add_object(sig_field)
            last_page["/Annots"].append(sig_field_ref)

            # Create AcroForm if it doesn't exist
            if "/AcroForm" not in writer._root_object:
                acro_form = DictionaryObject()
                acro_form.update({
                    NameObject("/Fields"): ArrayObject(),
                    NameObject("/SigFlags"): NumberObject(3),  # SignaturesExist | AppendOnly
                })
                writer._root_object[NameObject("/AcroForm")] = acro_form

            # Add field to AcroForm
            writer._root_object["/AcroForm"]["/Fields"].append(sig_field_ref)

        # Write the output
        with open(output_path, "wb") as output_file:
            writer.write(output_file)

    def _draw_form(self, c: canvas.Canvas, data: dict):
        """Draw the complete form on the canvas"""
        y = self.height - self.margin

        # Header with logo
        y = self._draw_header(c, y)

        # Date
        y = self._draw_date(c, y)

        # Transferred from section
        y = self._draw_transferred_from(c, y, data)

        # Assets table
        y = self._draw_assets_table(c, y, data.get("assets", []))

        # Transferred to section
        y = self._draw_transferred_to(c, y, data)

        # Declaration
        y = self._draw_declaration(c, y)

        # Signature
        self._draw_signature(c, y)

    def _draw_header(self, c: canvas.Canvas, y: float) -> float:
        """Draw the header with centered logo and titles"""
        logo_path = get_logo_path()
        if os.path.exists(logo_path):
            try:
                # Logo dimensions - maintain aspect ratio
                logo_height = 0.9 * inch
                logo_width = 2.8 * inch  # Approximate aspect ratio of the AU logo
                # Center the logo horizontally
                logo_x = (self.width - logo_width) / 2
                c.drawImage(logo_path, logo_x, y - logo_height,
                            width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        y -= 1.1 * inch

        # Main Store title
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(self.width / 2, y, "Main Store")

        y -= 0.5 * inch

        # ATF title
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(self.width / 2, y, "ATF")

        # Underline
        c.setStrokeColor(BLUE)
        c.line(self.width / 2 - 0.3 * inch, y - 0.05 * inch, self.width / 2 + 0.3 * inch, y - 0.05 * inch)

        y -= 0.3 * inch

        # Asset Transfer Form subtitle
        c.setFillColor(BLACK)
        c.setFont("Helvetica", 12)
        c.drawCentredString(self.width / 2, y, "(Asset Transfer Form)")

        # Underline
        c.line(self.width / 2 - 1 * inch, y - 0.05 * inch, self.width / 2 + 1 * inch, y - 0.05 * inch)

        return y - 0.5 * inch

    def _draw_date(self, c: canvas.Canvas, y: float) -> float:
        """Draw the date field"""
        date_text = get_form_date()

        c.setFillColor(BLACK)
        c.setStrokeColor(BLACK)
        c.setFont("Helvetica", 11)
        c.drawString(self.width - 2.5 * inch, y + 1.2 * inch, "Date :")

        # Date box - outline only
        c.rect(self.width - 2 * inch, y + 1.05 * inch, 1.3 * inch, 0.3 * inch, fill=0, stroke=1)
        c.drawString(self.width - 1.9 * inch, y + 1.15 * inch, date_text)

        return y

    def _truncate_text(self, c: canvas.Canvas, text: str, max_width: float, font_name: str, font_size: int) -> str:
        """Truncate text to fit within max_width"""
        if not text:
            return ""
        if c.stringWidth(text, font_name, font_size) <= max_width:
            return text
        # Truncate with ellipsis
        while len(text) > 0 and c.stringWidth(text + "...", font_name, font_size) > max_width:
            text = text[:-1]
        return text + "..." if text else ""

    def _draw_transferred_from(self, c: canvas.Canvas, y: float, data: dict) -> float:
        """Draw Transferred from section"""
        c.setFillColor(BLACK)
        c.setStrokeColor(BLACK)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(self.margin, y, "Transferred from:")

        # Underline
        c.line(self.margin, y - 0.05 * inch, self.margin + 1.2 * inch, y - 0.05 * inch)

        y -= 0.32 * inch

        # Custodian Name
        c.setFont("Helvetica", 10)
        c.drawString(self.margin, y, "Custodian Name:")
        name_x = self.margin + 1.35 * inch
        name_width = 3.5 * inch
        c.line(name_x, y - 0.05 * inch, name_x + name_width, y - 0.05 * inch)
        name_text = self._truncate_text(c, data.get("from_name", ""), name_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(name_x + 0.05 * inch, y, name_text)

        y -= 0.35 * inch

        # Department and Emp ID
        c.drawString(self.margin, y, "Department:")
        dept_x = self.margin + 1 * inch
        dept_width = 2.3 * inch
        c.line(dept_x, y - 0.05 * inch, dept_x + dept_width, y - 0.05 * inch)
        dept_text = self._truncate_text(c, data.get("from_department", ""), dept_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(dept_x + 0.05 * inch, y, dept_text)

        emp_label_x = 4.2 * inch
        c.drawString(emp_label_x, y, "Emp. ID:")
        emp_x = emp_label_x + 0.65 * inch
        emp_width = 1 * inch
        c.rect(emp_x, y - 0.08 * inch, emp_width, 0.28 * inch, fill=0, stroke=1)
        emp_text = self._truncate_text(c, data.get("from_emp_id", ""), emp_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(emp_x + 0.05 * inch, y, emp_text)

        return y - 0.4 * inch

    def _draw_assets_table(self, c: canvas.Canvas, y: float, assets: list) -> float:
        """Draw the assets table with dynamic rows"""
        # Table headers
        headers = ["No.", "Store Code", "Asset Name", "Description", "Old Asset No."]
        col_widths = [0.35 * inch, 1 * inch, 1.3 * inch, 2.3 * inch, 1.15 * inch]
        table_width = sum(col_widths)
        x_start = (self.width - table_width) / 2

        row_height = 0.3 * inch
        header_height = 0.35 * inch

        # Draw header row
        c.setStrokeColor(BLACK)
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 9)

        x = x_start
        for header, width in zip(headers, col_widths):
            c.rect(x, y - header_height, width, header_height)
            c.drawCentredString(x + width / 2, y - 0.22 * inch, header)
            x += width

        y -= header_height

        # Draw data rows - only as many as needed (no empty rows)
        num_rows = len(assets) if assets else 1
        c.setFont("Helvetica", 9)

        for row_num in range(num_rows):
            x = x_start
            for col, width in enumerate(col_widths):
                c.rect(x, y - row_height, width, row_height)

                if row_num < len(assets):
                    asset = assets[row_num]
                    if col == 0:
                        text = str(row_num + 1)
                    elif col == 1:
                        text = asset.get("store_code", "")
                    elif col == 2:
                        text = asset.get("asset_name", "")
                    elif col == 3:
                        text = asset.get("description", "")
                    elif col == 4:
                        text = asset.get("old_asset_no", "")

                    # Truncate text if too long
                    max_chars = int(width / 5.5)
                    if len(text) > max_chars and col != 0:
                        text = text[:max_chars]

                    if col == 0:
                        c.drawCentredString(x + width / 2, y - 0.19 * inch, text)
                    else:
                        c.drawString(x + 0.04 * inch, y - 0.19 * inch, text)

                x += width

            y -= row_height

        return y - 0.3 * inch

    def _draw_transferred_to(self, c: canvas.Canvas, y: float, data: dict) -> float:
        """Draw Transferred to section"""
        c.setFillColor(BLACK)
        c.setStrokeColor(BLACK)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(self.margin, y, "Transferred to:")

        # Underline
        c.line(self.margin, y - 0.05 * inch, self.margin + 1.1 * inch, y - 0.05 * inch)

        y -= 0.32 * inch

        # Custodian Name
        c.setFont("Helvetica", 10)
        c.drawString(self.margin, y, "Custodian Name:")
        name_x = self.margin + 1.35 * inch
        name_width = 3.5 * inch
        c.line(name_x, y - 0.05 * inch, name_x + name_width, y - 0.05 * inch)
        name_text = self._truncate_text(c, data.get("to_name", ""), name_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(name_x + 0.05 * inch, y, name_text)

        y -= 0.35 * inch

        # Department and Emp ID
        c.drawString(self.margin, y, "Department:")
        dept_x = self.margin + 1 * inch
        dept_width = 2.3 * inch
        c.line(dept_x, y - 0.05 * inch, dept_x + dept_width, y - 0.05 * inch)
        dept_text = self._truncate_text(c, data.get("to_department", ""), dept_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(dept_x + 0.05 * inch, y, dept_text)

        emp_label_x = 4.2 * inch
        c.drawString(emp_label_x, y, "Emp. ID:")
        emp_x = emp_label_x + 0.65 * inch
        emp_width = 1 * inch
        c.rect(emp_x, y - 0.08 * inch, emp_width, 0.28 * inch, fill=0, stroke=1)
        emp_text = self._truncate_text(c, data.get("to_emp_id", ""), emp_width - 0.1 * inch, "Helvetica", 10)
        c.drawString(emp_x + 0.05 * inch, y, emp_text)

        return y - 0.4 * inch

    def _draw_declaration(self, c: canvas.Canvas, y: float) -> float:
        """Draw the declaration text"""
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(self.margin, y, "Declaration:")

        # Underline
        c.line(self.margin, y - 0.05 * inch, self.margin + 0.9 * inch, y - 0.05 * inch)

        y -= 0.25 * inch
        c.setFont("Helvetica", 9)

        lines = [
            "This device is a property of AU and to be returned back to AU store after usage, this device can't",
            "be shifted to any other user without a written approval from the stores.",
            "I confirm that this device will be used for work purpose only.",
            "I also understand that I will be responsible for any misuse or damages that may occur."
        ]

        for line in lines:
            c.drawString(self.margin, y, line)
            y -= 0.16 * inch

        return y - 0.2 * inch

    def _draw_signature(self, c: canvas.Canvas, y: float):
        """Draw the signature section with certificate-based digital signature field"""
        # Check if we have enough space, if not create new page
        if y < 1.5 * inch:
            c.showPage()
            y = self.height - self.margin

        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(BLACK)
        c.drawString(self.margin, y, "Signature :")

        y -= 0.12 * inch

        # Signature box dimensions
        box_width = 2.2 * inch
        box_height = 0.9 * inch

        # Calculate coordinates for signature field (PDF coordinates)
        x1 = self.margin
        y1 = y - box_height
        x2 = self.margin + box_width
        y2 = y

        # Store coordinates for adding signature field later
        self.sig_field_rect = (x1, y1, x2, y2)

        # Draw signature box with transparent fill and thin border
        c.setStrokeColor(BLACK)
        c.setLineWidth(0.5)
        c.rect(x1, y1, box_width, box_height, fill=0, stroke=1)

        # Add helper text below the signature box
        c.setFont("Helvetica-Oblique", 8)
        c.setFillColor(colors.HexColor("#666666"))
        c.drawString(self.margin, y1 - 0.12 * inch, "Click here to sign in Adobe Acrobat")
