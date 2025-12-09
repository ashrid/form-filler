"""
Digital Signature Utilities
Generates text-based digital signatures for forms
"""

from datetime import datetime
import os
import sys


def get_base_path() -> str:
    """Returns the base path of the project (handles PyInstaller bundling)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_resource_path(relative_path: str) -> str:
    """Get path to bundled resource (for PyInstaller)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled - check _MEIPASS first, then executable dir
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            resource = os.path.join(meipass, relative_path)
            if os.path.exists(resource):
                return resource
        # Fall back to executable directory
        return os.path.join(os.path.dirname(sys.executable), relative_path)
    else:
        # Running as script
        return os.path.join(get_base_path(), relative_path)


class DigitalSignature:
    """Handles digital signature generation"""

    def __init__(self, name: str = "Rashid Ibrahim", timezone: str = "+04'00'"):
        self.name = name
        self.timezone = timezone

    def get_current_date(self) -> str:
        """Returns current date in DD/MM/YYYY format"""
        return datetime.now().strftime("%d/%m/%Y")

    def get_signature_date(self) -> str:
        """Returns date in signature format YYYY.MM.DD"""
        return datetime.now().strftime("%Y.%m.%d")

    def get_signature_time(self) -> str:
        """Returns time in signature format HH:MM:SS"""
        return datetime.now().strftime("%H:%M:%S")

    def get_full_signature_text(self) -> dict:
        """
        Returns signature data for PDF generation
        Format matches the Digital ID signature style
        """
        return {
            "name": self.name,
            "signed_by_text": f"Digitally signed by\n{self.name}",
            "date_text": f"Date: {self.get_signature_date()}",
            "time_text": f"{self.get_signature_time()} {self.timezone}"
        }


def get_form_date() -> str:
    """Returns current date for form header"""
    return datetime.now().strftime("%d/%m/%Y")


def get_logo_path() -> str:
    """Returns path to the Ajman University logo"""
    return get_resource_path(os.path.join("Forms", "Ajman-University-Logo.png"))


def get_output_path() -> str:
    """Returns path to output directory"""
    output_dir = os.path.join(get_base_path(), "output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir
