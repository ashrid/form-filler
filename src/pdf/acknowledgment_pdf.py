"""
Acknowledgment of Receipt Form PDF Generator
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
    IndirectObject,
)

from ..utils.signature import get_form_date, get_logo_path, get_output_path


# Colors matching the original form
BLUE = colors.HexColor("#0098DA")
BLACK = colors.black
WHITE = colors.white


class AcknowledgmentPDFGenerator:
    """Generates Acknowledgment of Receipt Form PDFs with digital signature field"""

    def __init__(self):
        self.width, self.height = A4
        self.margin = 0.6 * inch
        self.bottom_margin = 0.5 * inch
        self.sig_field_rect = None  # Will store signature field coordinates

    def _generate_filename(self, form_data: dict) -> str:
        """Generate filename in format: {emp ID} - {name} - acknowledgement form {asset name}.pdf"""
        emp_id = form_data.get("emp_id", "").strip() or "Unknown"
        name = form_data.get("custodian_name", "").strip() or "Unknown"

        # Get first asset name from items
        items = form_data.get("items", [])
        if items and items[0].get("description"):
            asset_name = items[0].get("description", "").strip()
            # Truncate asset name if too long
            if len(asset_name) > 30:
                asset_name = asset_name[:30]
        else:
            asset_name = "Asset"

        # Clean characters that aren't allowed in filenames
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            emp_id = emp_id.replace(char, '_')
            name = name.replace(char, '_')
            asset_name = asset_name.replace(char, '_')

        return f"{emp_id} - {name} - acknowledgement form {asset_name}.pdf"

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
                NameObject("/T"): TextStringObject("EmployeeSignature"),
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

    def _check_page_break(self, c: canvas.Canvas, y: float, needed_space: float) -> float:
        """Check if we need a page break and create new page if needed"""
        if y - needed_space < self.bottom_margin:
            c.showPage()
            return self.height - self.margin
        return y

    def _draw_form(self, c: canvas.Canvas, data: dict):
        """Draw the complete form on the canvas"""
        y = self.height - self.margin

        # Header with logo
        y = self._draw_header(c, y)

        # Date
        y = self._draw_date(c, y)

        # Items table
        y = self._draw_items_table(c, y, data.get("items", []))

        # Check if we need a new page for remaining content
        remaining_content_height = 4.5 * inch  # Approximate height needed for remaining sections
        y = self._check_page_break(c, y, remaining_content_height)

        # Custodian details
        y = self._draw_custodian_details(c, y, data)

        # Location section
        y = self._draw_location_section(c, y, data)

        # Declaration text
        y = self._draw_declaration(c, y)

        # Device type selection
        y = self._draw_device_selection(c, y, data)

        # Signature
        self._draw_signature(c, y, data)

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
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(self.width / 2, y, "Main Store")

        y -= 0.35 * inch

        # Form title
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(self.width / 2, y, "Acknowledgement of Receipt")

        return y - 0.4 * inch

    def _draw_date(self, c: canvas.Canvas, y: float) -> float:
        """Draw the date field"""
        date_text = get_form_date()

        c.setFillColor(BLACK)
        c.setFont("Helvetica", 11)
        c.drawString(self.width - 2.3 * inch, y + 0.25 * inch, "Date:")

        c.rect(self.width - 1.8 * inch, y + 0.1 * inch, 1.1 * inch, 0.28 * inch)
        c.drawString(self.width - 1.7 * inch, y + 0.18 * inch, date_text)

        return y

    def _draw_items_table(self, c: canvas.Canvas, y: float, items: list) -> float:
        """Draw the items table with dynamic rows"""
        headers = ["No.", "Store Code", "Item Description", "Qty.", "Purchase Date\n/LPO"]
        col_widths = [0.35 * inch, 1.1 * inch, 3.2 * inch, 0.45 * inch, 1.1 * inch]
        table_width = sum(col_widths)
        x_start = (self.width - table_width) / 2

        row_height = 0.32 * inch
        header_height = 0.4 * inch

        # Draw header row
        c.setStrokeColor(BLACK)
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 9)

        x = x_start
        for header, width in zip(headers, col_widths):
            c.rect(x, y - header_height, width, header_height)
            lines = header.split('\n')
            if len(lines) > 1:
                c.drawCentredString(x + width / 2, y - 0.15 * inch, lines[0])
                c.drawCentredString(x + width / 2, y - 0.28 * inch, lines[1])
            else:
                c.drawCentredString(x + width / 2, y - 0.25 * inch, header)
            x += width

        y -= header_height

        # Draw data rows - only as many as needed (no empty rows)
        num_rows = len(items) if items else 1
        c.setFont("Helvetica", 9)

        for row_num in range(num_rows):
            # Check for page break
            if y - row_height < self.bottom_margin + 4 * inch:
                c.showPage()
                y = self.height - self.margin
                # Redraw header on new page
                c.setStrokeColor(BLACK)
                c.setFillColor(BLACK)
                c.setFont("Helvetica-Bold", 9)
                x = x_start
                for header, width in zip(headers, col_widths):
                    c.rect(x, y - header_height, width, header_height)
                    lines = header.split('\n')
                    if len(lines) > 1:
                        c.drawCentredString(x + width / 2, y - 0.15 * inch, lines[0])
                        c.drawCentredString(x + width / 2, y - 0.28 * inch, lines[1])
                    else:
                        c.drawCentredString(x + width / 2, y - 0.25 * inch, header)
                    x += width
                y -= header_height
                c.setFont("Helvetica", 9)

            x = x_start
            for col, width in enumerate(col_widths):
                c.rect(x, y - row_height, width, row_height)

                if row_num < len(items):
                    item = items[row_num]
                    if col == 0:
                        text = str(row_num + 1)
                    elif col == 1:
                        text = item.get("store_code", "")
                    elif col == 2:
                        text = item.get("description", "")
                    elif col == 3:
                        text = item.get("qty", "")
                    elif col == 4:
                        text = item.get("purchase_date", "")

                    # Truncate if needed
                    max_chars = int(width / 5.5)
                    if len(text) > max_chars and col != 0:
                        text = text[:max_chars]

                    if col == 0:
                        c.drawCentredString(x + width / 2, y - 0.2 * inch, text)
                    else:
                        c.drawString(x + 0.04 * inch, y - 0.2 * inch, text)
                x += width

            y -= row_height

        return y - 0.25 * inch

    def _draw_custodian_details(self, c: canvas.Canvas, y: float, data: dict) -> float:
        """Draw custodian details section"""
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(self.margin, y, "Custodian Details:")

        y -= 0.28 * inch
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 9)

        c.drawString(self.margin, y, "Name:")
        c.line(self.margin + 0.45 * inch, y - 0.04 * inch, 4 * inch, y - 0.04 * inch)
        c.setFont("Helvetica", 9)
        c.drawString(self.margin + 0.5 * inch, y, data.get("custodian_name", ""))

        c.setFont("Helvetica-Bold", 9)
        c.drawString(4.3 * inch, y, "Emp. ID:")
        c.rect(4.9 * inch, y - 0.08 * inch, 0.9 * inch, 0.26 * inch)
        c.setFont("Helvetica", 9)
        c.drawString(4.95 * inch, y, data.get("emp_id", ""))

        y -= 0.35 * inch
        c.setFont("Helvetica-Bold", 9)
        c.drawString(self.margin, y, "College / Department:")
        c.line(self.margin + 1.35 * inch, y - 0.04 * inch, 5 * inch, y - 0.04 * inch)
        c.setFont("Helvetica", 9)
        c.drawString(self.margin + 1.4 * inch, y, data.get("department", ""))

        return y - 0.4 * inch

    def _draw_location_section(self, c: canvas.Canvas, y: float, data: dict) -> float:
        """Draw the location selection section"""
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 11)

        col1_x = self.margin
        col2_x = 2.6 * inch
        col3_x = 4.2 * inch

        c.drawString(col1_x, y, "Location: Building")
        c.drawString(col2_x, y, "Floor")
        c.drawString(col3_x, y, "Section")

        c.setStrokeColor(BLUE)
        c.line(col1_x, y - 0.04 * inch, col1_x + 1.35 * inch, y - 0.04 * inch)
        c.line(col2_x, y - 0.04 * inch, col2_x + 0.42 * inch, y - 0.04 * inch)
        c.line(col3_x, y - 0.04 * inch, col3_x + 0.55 * inch, y - 0.04 * inch)

        y -= 0.32 * inch
        c.setFillColor(BLACK)
        c.setFont("Helvetica", 9)

        buildings = ["SZH", "J1", "J2", "Student Hub", "Hostel", "Others:"]
        floors = ["Ground", "1st", "2nd", "3rd", "Others:"]
        sections = ["Male", "Female"]

        selected_building = data.get("building", "")
        selected_floor = data.get("floor", "")
        selected_section = data.get("section", "")

        line_height = 0.24 * inch

        for i, building in enumerate(buildings):
            by = y - i * line_height
            is_selected = building == selected_building or (building == "Others:" and selected_building == "Others")
            self._draw_radio(c, col1_x, by, is_selected)
            c.drawString(col1_x + 0.22 * inch, by, building)
            if building == "Others:" and selected_building == "Others":
                c.line(col1_x + 0.7 * inch, by - 0.04 * inch, col1_x + 1.6 * inch, by - 0.04 * inch)
                c.drawString(col1_x + 0.72 * inch, by, data.get("building_other", ""))

        for i, floor in enumerate(floors):
            fy = y - i * line_height
            is_selected = floor == selected_floor or (floor == "Others:" and selected_floor == "Others")
            self._draw_radio(c, col2_x, fy, is_selected)
            if floor in ["1st", "2nd", "3rd"]:
                c.drawString(col2_x + 0.22 * inch, fy, floor[0])
                c.setFont("Helvetica", 6)
                c.drawString(col2_x + 0.3 * inch, fy + 0.06 * inch, floor[1:])
                c.setFont("Helvetica", 9)
            else:
                c.drawString(col2_x + 0.22 * inch, fy, floor)
            if floor == "Others:" and selected_floor == "Others":
                c.line(col2_x + 0.65 * inch, fy - 0.04 * inch, col2_x + 1.3 * inch, fy - 0.04 * inch)
                c.drawString(col2_x + 0.67 * inch, fy, data.get("floor_other", ""))

        for i, section in enumerate(sections):
            sy = y - i * line_height
            self._draw_radio(c, col3_x, sy, section == selected_section)
            c.drawString(col3_x + 0.22 * inch, sy, section)

        return y - len(buildings) * line_height - 0.15 * inch

    def _draw_radio(self, c: canvas.Canvas, x: float, y: float, selected: bool):
        """Draw a radio button circle"""
        c.setStrokeColor(BLACK)
        c.circle(x + 0.08 * inch, y + 0.04 * inch, 0.065 * inch)
        if selected:
            c.setFillColor(BLACK)
            c.circle(x + 0.08 * inch, y + 0.04 * inch, 0.035 * inch, fill=1)

    def _draw_declaration(self, c: canvas.Canvas, y: float) -> float:
        """Draw the declaration text"""
        c.setFillColor(BLACK)
        c.setFont("Helvetica-BoldOblique", 9)

        text = ("I confirm that this device(s) is a property of Ajman University and to be "
                "returned back to AU Store after usage. This device(s) can't be shifted to "
                "any other user/location without a written approval from the Store.")

        words = text.split()
        lines = []
        current_line = []
        max_width = self.width - 2 * self.margin

        for word in words:
            current_line.append(word)
            test_line = " ".join(current_line)
            if c.stringWidth(test_line, "Helvetica-BoldOblique", 9) > max_width:
                current_line.pop()
                lines.append(" ".join(current_line))
                current_line = [word]
        if current_line:
            lines.append(" ".join(current_line))

        for line in lines:
            c.drawString(self.margin, y, line)
            y -= 0.16 * inch

        return y - 0.12 * inch

    def _draw_device_selection(self, c: canvas.Canvas, y: float, data: dict) -> float:
        """Draw device type selection"""
        c.setFont("Helvetica-BoldOblique", 9)
        c.setFillColor(BLACK)
        c.drawString(self.margin, y, "Please select one of the following:")

        y -= 0.28 * inch
        device_type = data.get("device_type", "")

        # Office device
        self._draw_radio(c, self.margin + 0.15 * inch, y, device_type == "Office")
        c.setFont("Helvetica-Bold", 9)
        c.drawString(self.margin + 0.4 * inch, y, "Office Device")
        y -= 0.18 * inch
        c.setFont("Helvetica", 8)
        office_text = ("I understand that I will be responsible for any misuse or damages that may occur. "
                       "I confirm that this device(s) will be used for work purpose only.")

        # Word wrap for office text
        words = office_text.split()
        lines = []
        current_line = []
        max_width = self.width - 2 * self.margin - 0.4 * inch

        for word in words:
            current_line.append(word)
            test_line = " ".join(current_line)
            if c.stringWidth(test_line, "Helvetica", 8) > max_width:
                current_line.pop()
                lines.append(" ".join(current_line))
                current_line = [word]
        if current_line:
            lines.append(" ".join(current_line))

        for line in lines:
            c.drawString(self.margin + 0.4 * inch, y, line)
            y -= 0.13 * inch

        y -= 0.12 * inch

        # Lab device
        self._draw_radio(c, self.margin + 0.15 * inch, y, device_type == "Lab")
        c.setFont("Helvetica-Bold", 9)
        c.drawString(self.margin + 0.4 * inch, y, "Lab Device")
        y -= 0.18 * inch
        c.setFont("Helvetica", 8)
        c.drawString(self.margin + 0.4 * inch, y,
                     "I understand that the lab supervisor shall monitor the lab devices to avoid any misuse or damage.")

        return y - 0.3 * inch

    def _draw_signature(self, c: canvas.Canvas, y: float, data: dict):
        """Draw the signature section with certificate-based digital signature field"""
        # Check if we have enough space, if not create new page
        if y < 1.3 * inch:
            c.showPage()
            y = self.height - self.margin

        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(BLACK)
        c.drawString(self.margin, y, "Employee Signature:")

        y -= 0.12 * inch

        # Signature box dimensions
        box_width = 2.2 * inch
        box_height = 1 * inch

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
