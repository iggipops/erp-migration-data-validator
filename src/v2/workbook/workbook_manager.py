from openpyxl.workbook import Workbook

from config.config_loader import AppConfig
from workbook.excel_utils import copy_workbook, open_workbook, save_workbook
from utils.logger import get_logger


class WorkbookManager:
    """
    Handles copy/open/save of the main input/output workbook.
    External file import is handled separately by importer.Importer
    (FS section 8.5, architectural split from V2.3.r2+).
    """

    def __init__(self, config: AppConfig):
        self.config   = config
        self.logger   = get_logger()
        self.workbook: Workbook | None = None

    # -----------------------------------------------------------------------
    # Section 8.3 — Copy input to output
    # -----------------------------------------------------------------------

    def copy_and_open(self) -> Workbook:
        """
        Copy input workbook to output path and open it.
        All subsequent operations apply to the output file only.
        """
        self.logger.info(
            f"Copying input workbook: "
            f"{self.config.input_file} -> {self.config.output_file}"
        )
        copy_workbook(self.config.input_file, self.config.output_file)
        self.workbook = open_workbook(self.config.output_file)
        self.logger.info("Output workbook opened successfully")
        return self.workbook

    def save(self) -> None:
        if self.workbook is None:
            raise RuntimeError("Workbook is not open")
        save_workbook(self.workbook, self.config.output_file)
        self.logger.info(f"Output workbook saved: {self.config.output_file}")
