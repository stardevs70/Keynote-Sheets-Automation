"""
Google Sheets API Client Module
Handles authentication and data fetching from Google Sheets.
"""

import os
import logging
from dataclasses import dataclass
from typing import Any, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Sheets API scope - read-only access
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

logger = logging.getLogger(__name__)


@dataclass
class MappingRow:
    """Represents a single mapping row from KeynoteMap sheet."""
    id: str
    sheet_range: str
    slide_index: int
    target_type: str  # 'shape' or 'table_cell'
    object_name: str
    row: Optional[int] = None
    col: Optional[int] = None
    format: str = 'text'
    prefix: str = ''
    suffix: str = ''
    notes: str = ''

    @classmethod
    def from_row(cls, row: list) -> 'MappingRow':
        """Create a MappingRow from a spreadsheet row (list of cell values)."""
        def get_cell(index: int, default: Any = '') -> Any:
            try:
                return row[index] if row[index] is not None else default
            except IndexError:
                return default

        def get_int(index: int, default: Optional[int] = None) -> Optional[int]:
            try:
                val = row[index]
                if val is None or val == '':
                    return default
                return int(float(val))
            except (IndexError, ValueError, TypeError):
                return default

        return cls(
            id=str(get_cell(0, '')),
            sheet_range=str(get_cell(1, '')),
            slide_index=get_int(2, 1) or 1,
            target_type=str(get_cell(3, 'shape')).lower(),
            object_name=str(get_cell(4, '')),
            row=get_int(5),
            col=get_int(6),
            format=str(get_cell(7, 'text')),
            prefix=str(get_cell(8, '')),
            suffix=str(get_cell(9, '')),
            notes=str(get_cell(10, ''))
        )


class SheetsClient:
    """Client for interacting with Google Sheets API."""

    def __init__(self, config: dict):
        """
        Initialize the Sheets client.

        Args:
            config: Dictionary containing Google API configuration with keys:
                - credentials_file: Path to OAuth credentials JSON
                - token_file: Path to store/retrieve OAuth tokens
                - spreadsheet_id: ID of the Google Spreadsheet
                - mapping_sheet: Name of the mapping tab
        """
        self.credentials_file = config.get('credentials_file', 'credentials.json')
        self.token_file = config.get('token_file', 'token.json')
        self.spreadsheet_id = config.get('spreadsheet_id', '')
        self.mapping_sheet = config.get('mapping_sheet', 'KeynoteMap')
        self._service = None
        self._creds = None

    def _get_credentials(self) -> Credentials:
        """Get or refresh OAuth credentials."""
        creds = None

        # Load existing token if available
        if os.path.exists(self.token_file):
            try:
                creds = Credentials.from_authorized_user_file(self.token_file, SCOPES)
                logger.debug("Loaded existing credentials from token file")
            except Exception as e:
                logger.warning(f"Failed to load token file: {e}")

        # If no valid credentials, initiate OAuth flow
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    logger.info("Refreshed expired credentials")
                except Exception as e:
                    logger.warning(f"Failed to refresh credentials: {e}")
                    creds = None

            if not creds:
                if not os.path.exists(self.credentials_file):
                    raise FileNotFoundError(
                        f"Credentials file '{self.credentials_file}' not found. "
                        "Please download it from Google Cloud Console."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES
                )
                creds = flow.run_local_server(port=0)
                logger.info("Completed OAuth flow for new credentials")

            # Save credentials for next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
                logger.debug(f"Saved credentials to {self.token_file}")

        return creds

    def _get_service(self):
        """Get or create the Sheets API service."""
        if self._service is None:
            self._creds = self._get_credentials()
            self._service = build('sheets', 'v4', credentials=self._creds)
        return self._service

    def read_mapping(self) -> list[MappingRow]:
        """
        Read the mapping configuration from the KeynoteMap sheet.

        Returns:
            List of MappingRow objects representing each mapping.
        """
        try:
            service = self._get_service()
            sheet = service.spreadsheets()

            # Read all data from the mapping sheet (skip header row)
            range_name = f"'{self.mapping_sheet}'!A2:K"
            result = sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])
            if not values:
                logger.warning(f"No mapping data found in {self.mapping_sheet}")
                return []

            mappings = []
            for i, row in enumerate(values, start=2):
                # Skip empty rows
                if not row or not any(cell for cell in row):
                    continue
                # Skip rows without sheet_range (column B)
                if len(row) < 2 or not row[1]:
                    continue

                try:
                    mapping = MappingRow.from_row(row)
                    mappings.append(mapping)
                    logger.debug(f"Loaded mapping: {mapping.id} -> {mapping.object_name}")
                except Exception as e:
                    logger.warning(f"Failed to parse mapping row {i}: {e}")

            logger.info(f"Loaded {len(mappings)} mappings from {self.mapping_sheet}")
            return mappings

        except HttpError as e:
            logger.error(f"Google Sheets API error: {e}")
            raise
        except Exception as e:
            logger.error(f"Failed to read mapping: {e}")
            raise

    def batch_get_values(self, ranges: list[str]) -> dict[str, Any]:
        """
        Fetch multiple cell values in a single API call.

        Args:
            ranges: List of A1 notation ranges (e.g., ["Data Vault!B12", "Data Vault!C15"])

        Returns:
            Dictionary mapping range string to cell value.
        """
        if not ranges:
            return {}

        try:
            service = self._get_service()
            sheet = service.spreadsheets()

            result = sheet.values().batchGet(
                spreadsheetId=self.spreadsheet_id,
                ranges=ranges
            ).execute()

            value_ranges = result.get('valueRanges', [])
            values_dict = {}

            for i, vr in enumerate(value_ranges):
                range_key = ranges[i]
                values = vr.get('values', [[]])
                # Get the first cell value (most mappings are single cells)
                if values and values[0]:
                    values_dict[range_key] = values[0][0]
                else:
                    values_dict[range_key] = None
                logger.debug(f"Fetched {range_key}: {values_dict[range_key]}")

            logger.info(f"Fetched {len(values_dict)} values from spreadsheet")
            return values_dict

        except HttpError as e:
            logger.error(f"Google Sheets API error during batch get: {e}")
            raise
        except Exception as e:
            logger.error(f"Failed to batch get values: {e}")
            raise

    def get_single_value(self, range_name: str) -> Any:
        """
        Fetch a single cell value.

        Args:
            range_name: A1 notation range (e.g., "Data Vault!B12")

        Returns:
            The cell value, or None if empty.
        """
        result = self.batch_get_values([range_name])
        return result.get(range_name)


class MockSheetsClient:
    """Mock client for testing without Google API credentials."""

    def __init__(self, config: dict):
        self.mapping_sheet = config.get('mapping_sheet', 'KeynoteMap')
        self._mock_mappings = config.get('mock_mappings', [])
        self._mock_values = config.get('mock_values', {})

    def read_mapping(self) -> list[MappingRow]:
        """Return mock mappings."""
        logger.info(f"[MOCK] Loaded {len(self._mock_mappings)} mappings")
        return self._mock_mappings

    def batch_get_values(self, ranges: list[str]) -> dict[str, Any]:
        """Return mock values."""
        result = {}
        for r in ranges:
            result[r] = self._mock_values.get(r, f"MockValue_{r}")
        logger.info(f"[MOCK] Fetched {len(result)} values")
        return result

    def get_single_value(self, range_name: str) -> Any:
        """Return a mock value."""
        return self._mock_values.get(range_name)


def create_client(config: dict) -> SheetsClient:
    """Factory function to create a SheetsClient from config."""
    google_config = config.get('google', {})

    # Use mock client if mock mode is enabled
    if config.get('mock_mode', False):
        logger.info("Using MockSheetsClient for testing")
        return MockSheetsClient(google_config)

    return SheetsClient(google_config)
