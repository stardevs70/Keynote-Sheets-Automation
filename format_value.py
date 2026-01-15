"""
Value Formatting Module
Converts raw values from Google Sheets into formatted strings for Keynote display.
"""

import logging
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any, Optional

logger = logging.getLogger(__name__)

# Google Sheets epoch (dates are stored as days since this date)
SHEETS_EPOCH = datetime(1899, 12, 30)


def parse_number(value: Any) -> Optional[float]:
    """
    Parse a value into a float, handling various input formats.

    Args:
        value: The raw value (could be string, int, float, etc.)

    Returns:
        Float value or None if parsing fails.
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):
        # Remove common formatting characters
        cleaned = value.strip()
        cleaned = cleaned.replace(',', '')
        cleaned = cleaned.replace('$', '')
        cleaned = cleaned.replace('%', '')
        cleaned = cleaned.replace(' ', '')

        # Handle parentheses for negative numbers
        if cleaned.startswith('(') and cleaned.endswith(')'):
            cleaned = '-' + cleaned[1:-1]

        try:
            return float(cleaned)
        except ValueError:
            return None

    return None


def format_currency(value: float, decimals: int = 0, symbol: str = '$') -> str:
    """
    Format a number as currency.

    Args:
        value: The numeric value
        decimals: Number of decimal places (0, 1, or 2)
        symbol: Currency symbol (default: $)

    Returns:
        Formatted currency string (e.g., "$5,000" or "$5,000.00")
    """
    if decimals == 0:
        formatted = f"{abs(value):,.0f}"
    elif decimals == 1:
        formatted = f"{abs(value):,.1f}"
    else:
        formatted = f"{abs(value):,.2f}"

    if value < 0:
        return f"-{symbol}{formatted}"
    return f"{symbol}{formatted}"


def format_percent(value: float, decimals: int = 1, multiply: bool = False) -> str:
    """
    Format a number as a percentage.

    Args:
        value: The numeric value
        decimals: Number of decimal places
        multiply: If True, multiply by 100 (for values like 0.133 -> 13.3%)

    Returns:
        Formatted percentage string (e.g., "13.3%")
    """
    if multiply:
        value = value * 100

    if decimals == 0:
        return f"{value:.0f}%"
    elif decimals == 1:
        return f"{value:.1f}%"
    else:
        return f"{value:.{decimals}f}%"


def format_integer(value: float) -> str:
    """
    Format a number as an integer with thousand separators.

    Args:
        value: The numeric value

    Returns:
        Formatted integer string (e.g., "10,000")
    """
    return f"{int(round(value)):,}"


def format_decimal(value: float, decimals: int = 2) -> str:
    """
    Format a number with specified decimal places.

    Args:
        value: The numeric value
        decimals: Number of decimal places

    Returns:
        Formatted decimal string (e.g., "5.00")
    """
    return f"{value:,.{decimals}f}"


def format_date_mdy(value: Any) -> str:
    """
    Format a date value in "Month Day, Year" format.

    Args:
        value: The date value (can be serial number, datetime, or string)

    Returns:
        Formatted date string (e.g., "April 1, 2021")
    """
    date_obj = None

    if isinstance(value, datetime):
        date_obj = value
    elif isinstance(value, (int, float)):
        # Google Sheets stores dates as serial numbers (days since epoch)
        try:
            date_obj = SHEETS_EPOCH + timedelta(days=int(value))
        except (ValueError, OverflowError):
            pass
    elif isinstance(value, str):
        # Try common date formats
        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%B %d, %Y']:
            try:
                date_obj = datetime.strptime(value.strip(), fmt)
                break
            except ValueError:
                continue

    if date_obj:
        # Use platform-independent formatting (%-d doesn't work on Windows)
        month_name = date_obj.strftime('%B')
        return f"{month_name} {date_obj.day}, {date_obj.year}"

    # Fallback: return as-is if we can't parse it
    return str(value)


def format_date_short(value: Any) -> str:
    """
    Format a date value in "M/YYYY" format.

    Args:
        value: The date value

    Returns:
        Formatted date string (e.g., "1/2027")
    """
    date_obj = None

    if isinstance(value, datetime):
        date_obj = value
    elif isinstance(value, (int, float)):
        try:
            date_obj = SHEETS_EPOCH + timedelta(days=int(value))
        except (ValueError, OverflowError):
            pass
    elif isinstance(value, str):
        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
            try:
                date_obj = datetime.strptime(value.strip(), fmt)
                break
            except ValueError:
                continue

    if date_obj:
        return f"{date_obj.month}/{date_obj.year}"

    return str(value)


def format_text_number(value: Any) -> str:
    """
    Convert numbers to their word equivalents (for counts like "Ten").

    Args:
        value: The numeric value

    Returns:
        Number as word (e.g., 10 -> "Ten")
    """
    number_words = {
        0: 'Zero', 1: 'One', 2: 'Two', 3: 'Three', 4: 'Four',
        5: 'Five', 6: 'Six', 7: 'Seven', 8: 'Eight', 9: 'Nine',
        10: 'Ten', 11: 'Eleven', 12: 'Twelve', 13: 'Thirteen',
        14: 'Fourteen', 15: 'Fifteen', 16: 'Sixteen', 17: 'Seventeen',
        18: 'Eighteen', 19: 'Nineteen', 20: 'Twenty'
    }

    num = parse_number(value)
    if num is not None:
        int_val = int(round(num))
        if int_val in number_words:
            return number_words[int_val]

    return str(value)


def format_value(raw_value: Any, fmt: str, prefix: str = '', suffix: str = '',
                 empty_value: str = '') -> str:
    """
    Format a raw value according to the specified format type.

    Args:
        raw_value: The raw value from Google Sheets
        fmt: Format type ('currency0', 'currency2', 'percent1', 'integer',
             'decimal2', 'date_mdy', 'date_short', 'text_number', 'text')
        prefix: Text to prepend to the result
        suffix: Text to append to the result
        empty_value: Value to return if raw_value is empty/None

    Returns:
        Formatted string ready for Keynote display.
    """
    # Handle empty values
    if raw_value is None or (isinstance(raw_value, str) and raw_value.strip() == ''):
        return empty_value

    fmt = fmt.lower().strip() if fmt else 'text'
    result = ''

    try:
        if fmt == 'currency0':
            num = parse_number(raw_value)
            if num is not None:
                result = format_currency(num, decimals=0)
            else:
                result = str(raw_value)

        elif fmt == 'currency1':
            num = parse_number(raw_value)
            if num is not None:
                result = format_currency(num, decimals=1)
            else:
                result = str(raw_value)

        elif fmt == 'currency2':
            num = parse_number(raw_value)
            if num is not None:
                result = format_currency(num, decimals=2)
            else:
                result = str(raw_value)

        elif fmt == 'percent0':
            num = parse_number(raw_value)
            if num is not None:
                # Check if value is already in percentage form (> 1) or decimal (< 1)
                multiply = abs(num) <= 1 and abs(num) > 0
                result = format_percent(num, decimals=0, multiply=multiply)
            else:
                result = str(raw_value)

        elif fmt == 'percent1':
            num = parse_number(raw_value)
            if num is not None:
                multiply = abs(num) <= 1 and abs(num) > 0
                result = format_percent(num, decimals=1, multiply=multiply)
            else:
                result = str(raw_value)

        elif fmt == 'percent2':
            num = parse_number(raw_value)
            if num is not None:
                multiply = abs(num) <= 1 and abs(num) > 0
                result = format_percent(num, decimals=2, multiply=multiply)
            else:
                result = str(raw_value)

        elif fmt == 'integer':
            num = parse_number(raw_value)
            if num is not None:
                result = format_integer(num)
            else:
                result = str(raw_value)

        elif fmt == 'decimal1':
            num = parse_number(raw_value)
            if num is not None:
                result = format_decimal(num, decimals=1)
            else:
                result = str(raw_value)

        elif fmt == 'decimal2':
            num = parse_number(raw_value)
            if num is not None:
                result = format_decimal(num, decimals=2)
            else:
                result = str(raw_value)

        elif fmt == 'date_mdy':
            result = format_date_mdy(raw_value)

        elif fmt == 'date_short':
            result = format_date_short(raw_value)

        elif fmt == 'text_number':
            result = format_text_number(raw_value)

        else:  # 'text' or unknown format
            result = str(raw_value)

    except Exception as e:
        logger.warning(f"Error formatting value '{raw_value}' with format '{fmt}': {e}")
        result = str(raw_value)

    # Apply prefix and suffix
    # Note: For currency formats, the $ is already included, so prefix should be empty
    # unless the user wants additional text
    final_result = f"{prefix}{result}{suffix}"

    logger.debug(f"Formatted '{raw_value}' with '{fmt}' -> '{final_result}'")
    return final_result


# Format type descriptions for documentation
FORMAT_TYPES = {
    'currency0': 'Currency with no decimals (e.g., "$5,000")',
    'currency1': 'Currency with 1 decimal (e.g., "$5,000.0")',
    'currency2': 'Currency with 2 decimals (e.g., "$5,000.00")',
    'percent0': 'Percentage with no decimals (e.g., "13%")',
    'percent1': 'Percentage with 1 decimal (e.g., "13.3%")',
    'percent2': 'Percentage with 2 decimals (e.g., "13.30%")',
    'integer': 'Whole number with commas (e.g., "10,000")',
    'decimal1': 'Number with 1 decimal (e.g., "5.0")',
    'decimal2': 'Number with 2 decimals (e.g., "5.00")',
    'date_mdy': 'Full date format (e.g., "April 1, 2021")',
    'date_short': 'Short date format (e.g., "1/2027")',
    'text_number': 'Number as word (e.g., "Ten")',
    'text': 'Plain text (no formatting)',
}
