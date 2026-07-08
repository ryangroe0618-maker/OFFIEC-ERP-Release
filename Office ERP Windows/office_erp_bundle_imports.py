# -*- coding: utf-8 -*-
"""Imports used only so PyInstaller bundles ERP runtime dependencies."""

import gspread  # noqa: F401
import numpy  # noqa: F401
import openpyxl  # noqa: F401
import pandas  # noqa: F401
import requests  # noqa: F401
from google.auth.exceptions import GoogleAuthError  # noqa: F401
from google.oauth2 import service_account  # noqa: F401
from openpyxl.cell.rich_text import CellRichText, TextBlock  # noqa: F401
from openpyxl.cell.text import InlineFont  # noqa: F401
from openpyxl.chart import BarChart, Reference  # noqa: F401
from openpyxl.chart.label import DataLabelList  # noqa: F401
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor  # noqa: F401
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: F401
from openpyxl.utils import get_column_letter  # noqa: F401
from pandas.errors import EmptyDataError  # noqa: F401
from pypdf import PdfReader, PdfWriter  # noqa: F401
