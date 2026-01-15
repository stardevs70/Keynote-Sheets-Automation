#!/usr/bin/env python3
"""
Full integration test for the PowerPoint update script.
Tests the complete flow from mapping to presentation update with mock data.
"""

import logging
import sys
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

from sheets_client import MappingRow, MockSheetsClient
from format_value import format_value
from powerpoint_bridge import PowerPointBridge


def create_test_config():
    """Create a test configuration with mock data."""

    # Define mock mappings (matching our sample_investor_deck.pptx)
    mock_mappings = [
        # Slide 1: Title slide
        MappingRow(
            id="report_date",
            sheet_range="DataVault!B2",
            slide_index=1,
            target_type="shape",
            object_name="ReportDate",
            format="text",
            prefix="",
            suffix=""
        ),
        # Slide 2: Key Metrics
        MappingRow(
            id="revenue_value",
            sheet_range="DataVault!B5",
            slide_index=2,
            target_type="shape",
            object_name="RevenueValue",
            format="currency0",
            prefix="",
            suffix=""
        ),
        MappingRow(
            id="growth_value",
            sheet_range="DataVault!B6",
            slide_index=2,
            target_type="shape",
            object_name="GrowthValue",
            format="percent1",
            prefix="",
            suffix=""
        ),
        MappingRow(
            id="customer_value",
            sheet_range="DataVault!B7",
            slide_index=2,
            target_type="shape",
            object_name="CustomerValue",
            format="integer",
            prefix="",
            suffix=""
        ),
        # Slide 3: Financial Table
        MappingRow(
            id="q4_revenue",
            sheet_range="DataVault!D10",
            slide_index=3,
            target_type="table_cell",
            object_name="FinancialTable",
            row=2,
            col=4,
            format="currency0"
        ),
        MappingRow(
            id="q4_expenses",
            sheet_range="DataVault!D11",
            slide_index=3,
            target_type="table_cell",
            object_name="FinancialTable",
            row=3,
            col=4,
            format="currency0"
        ),
        MappingRow(
            id="q4_net_income",
            sheet_range="DataVault!D12",
            slide_index=3,
            target_type="table_cell",
            object_name="FinancialTable",
            row=4,
            col=4,
            format="currency0"
        ),
        # Slide 4: KPI Table
        MappingRow(
            id="cac_current",
            sheet_range="DataVault!E15",
            slide_index=4,
            target_type="table_cell",
            object_name="KPITable",
            row=2,
            col=2,
            format="currency0"
        ),
        MappingRow(
            id="ltv_current",
            sheet_range="DataVault!E16",
            slide_index=4,
            target_type="table_cell",
            object_name="KPITable",
            row=3,
            col=2,
            format="currency0"
        ),
        MappingRow(
            id="churn_current",
            sheet_range="DataVault!E17",
            slide_index=4,
            target_type="table_cell",
            object_name="KPITable",
            row=4,
            col=2,
            format="percent1"
        ),
        # Slide 5: Summary
        MappingRow(
            id="total_revenue_summary",
            sheet_range="DataVault!B20",
            slide_index=5,
            target_type="shape",
            object_name="TotalRevenue",
            format="currency0",
            prefix="Total Revenue: ",
            suffix=""
        ),
        MappingRow(
            id="yoy_growth_summary",
            sheet_range="DataVault!B21",
            slide_index=5,
            target_type="shape",
            object_name="YoYGrowth",
            format="percent1",
            prefix="Year-over-Year Growth: ",
            suffix=""
        ),
        MappingRow(
            id="inception_date",
            sheet_range="DataVault!B22",
            slide_index=5,
            target_type="shape",
            object_name="InceptionDate",
            format="date_mdy",
            prefix="Fund Inception: ",
            suffix=""
        ),
    ]

    # Define mock values (simulating Google Sheets data)
    mock_values = {
        "DataVault!B2": "Q1 2025",
        "DataVault!B5": 7500000,
        "DataVault!B6": 0.325,
        "DataVault!B7": 1850,
        "DataVault!D10": 7500000,
        "DataVault!D11": 4200000,
        "DataVault!D12": 3300000,
        "DataVault!E15": 125,
        "DataVault!E16": 3200,
        "DataVault!E17": 0.042,
        "DataVault!B20": 7500000,
        "DataVault!B21": 0.325,
        "DataVault!B22": 44287,  # Excel/Sheets serial date for April 1, 2021
    }

    return {
        'mock_mode': True,
        'google': {
            'spreadsheet_id': 'TEST_SPREADSHEET_ID',
            'mapping_sheet': 'KeynoteMap',
            'mock_mappings': mock_mappings,
            'mock_values': mock_values,
        },
        'powerpoint': {
            'file_path': str(script_dir / 'sample_investor_deck.pptx')
        },
        'defaults': {
            'empty_value': '',
            'error_value': '#ERROR'
        }
    }


def process_mapping(bridge: PowerPointBridge, mapping: MappingRow, value, config: dict,
                    dry_run: bool = False) -> tuple:
    """Process a single mapping (copied from update_presentation.py for testing)."""
    defaults = config.get('defaults', {})
    empty_value = defaults.get('empty_value', '')

    formatted_text = format_value(
        raw_value=value,
        fmt=mapping.format,
        prefix=mapping.prefix,
        suffix=mapping.suffix,
        empty_value=empty_value
    )

    logger.debug(f"Mapping '{mapping.id}': {value} -> '{formatted_text}'")

    if dry_run:
        if mapping.target_type == 'shape':
            logger.info(f"[DRY RUN] Would update shape '{mapping.object_name}' on slide {mapping.slide_index} to: {formatted_text}")
        else:
            logger.info(f"[DRY RUN] Would update table '{mapping.object_name}' cell ({mapping.row},{mapping.col}) on slide {mapping.slide_index} to: {formatted_text}")
        return True, "Dry run - no changes made"

    if mapping.target_type == 'shape':
        return bridge.update_shape_text(
            slide_index=mapping.slide_index,
            shape_name=mapping.object_name,
            new_text=formatted_text
        )
    elif mapping.target_type == 'table_cell':
        if mapping.row is None or mapping.col is None:
            return False, f"Table cell mapping '{mapping.id}' is missing row or col index"
        return bridge.update_table_cell(
            slide_index=mapping.slide_index,
            table_name=mapping.object_name,
            row=mapping.row,
            col=mapping.col,
            new_text=formatted_text
        )
    else:
        return False, f"Unknown target type '{mapping.target_type}' for mapping '{mapping.id}'"


def test_dry_run():
    """Test the full update flow in dry-run mode."""
    print("\n" + "=" * 60)
    print("TEST: Full Update Flow (DRY RUN)")
    print("=" * 60)

    config = create_test_config()

    # Create mock sheets client
    sheets_client = MockSheetsClient(config['google'])

    # Get mappings and values
    mappings = sheets_client.read_mapping()
    print(f"Loaded {len(mappings)} mappings")

    ranges = sorted(set(m.sheet_range for m in mappings if m.sheet_range))
    values_by_range = sheets_client.batch_get_values(ranges)
    print(f"Fetched {len(values_by_range)} values")

    # Open presentation
    pptx_path = config['powerpoint']['file_path']
    bridge = PowerPointBridge(pptx_path)
    assert bridge.open(), f"Failed to open {pptx_path}"
    print(f"Opened presentation: {pptx_path}")

    # Process each mapping in dry-run mode
    success_count = 0
    error_count = 0

    print("\n--- DRY RUN UPDATES ---")
    for mapping in mappings:
        raw_value = values_by_range.get(mapping.sheet_range)
        success, message = process_mapping(bridge, mapping, raw_value, config, dry_run=True)
        if success:
            success_count += 1
        else:
            error_count += 1
            print(f"ERROR: {message}")

    print(f"\nDry run complete: {success_count} successful, {error_count} errors")
    assert error_count == 0, f"Encountered {error_count} errors during dry run"
    print("\nPASS: Full Update Flow (DRY RUN)")


def test_actual_update():
    """Test the full update flow with actual changes."""
    print("\n" + "=" * 60)
    print("TEST: Full Update Flow (ACTUAL)")
    print("=" * 60)

    config = create_test_config()

    # Create mock sheets client
    sheets_client = MockSheetsClient(config['google'])

    # Get mappings and values
    mappings = sheets_client.read_mapping()
    ranges = sorted(set(m.sheet_range for m in mappings if m.sheet_range))
    values_by_range = sheets_client.batch_get_values(ranges)

    # Open presentation
    pptx_path = config['powerpoint']['file_path']
    bridge = PowerPointBridge(pptx_path)
    assert bridge.open(), f"Failed to open {pptx_path}"

    # Process each mapping with actual updates
    success_count = 0
    error_count = 0
    errors = []

    print("\n--- ACTUAL UPDATES ---")
    for mapping in mappings:
        raw_value = values_by_range.get(mapping.sheet_range)
        success, message = process_mapping(bridge, mapping, raw_value, config, dry_run=False)
        if success:
            success_count += 1
            print(f"OK: {mapping.id} -> {message}")
        else:
            error_count += 1
            errors.append(f"{mapping.id}: {message}")
            print(f"FAIL: {mapping.id} -> {message}")

    print(f"\nUpdate complete: {success_count} successful, {error_count} errors")

    # Save to test output
    output_path = script_dir / "test_updated_deck.pptx"
    assert bridge.save(str(output_path)), "Failed to save presentation"
    print(f"Saved to: {output_path}")

    # Verify the output file
    assert output_path.exists(), "Output file not created"
    print(f"Output file size: {output_path.stat().st_size} bytes")

    # Clean up
    output_path.unlink()
    print(f"Cleaned up: {output_path}")

    assert error_count == 0, f"Encountered {error_count} errors: {errors}"
    print("\nPASS: Full Update Flow (ACTUAL)")


def test_format_value_integration():
    """Test format_value function with various inputs."""
    print("\n" + "=" * 60)
    print("TEST: Format Value Integration")
    print("=" * 60)

    test_cases = [
        # (raw_value, format, prefix, suffix, expected_contains)
        (7500000, "currency0", "", "", "$7,500,000"),
        (0.325, "percent1", "", "", "32.5%"),
        (1850, "integer", "", "", "1,850"),
        (125, "currency0", "", "", "$125"),
        (44287, "date_mdy", "", "", "April 1, 2021"),
        ("Q1 2025", "text", "", "", "Q1 2025"),
        (7500000, "currency0", "Total Revenue: ", "", "Total Revenue: $7,500,000"),
        (0.325, "percent1", "Growth: ", " YoY", "Growth: 32.5% YoY"),
    ]

    for raw_value, fmt, prefix, suffix, expected in test_cases:
        result = format_value(raw_value, fmt, prefix, suffix)
        print(f"  format_value({raw_value}, '{fmt}', '{prefix}', '{suffix}') = '{result}'")
        assert expected in result, f"Expected '{expected}' in '{result}'"

    print("\nPASS: Format Value Integration")


def run_all_tests():
    """Run all integration tests."""
    print("\n" + "=" * 60)
    print("FULL INTEGRATION TESTS")
    print("=" * 60)

    try:
        test_format_value_integration()
        test_dry_run()
        test_actual_update()

        print("\n" + "=" * 60)
        print("ALL INTEGRATION TESTS PASSED!")
        print("=" * 60)
        return 0

    except AssertionError as e:
        print(f"\nTEST FAILED: {e}")
        return 1
    except Exception as e:
        print(f"\nTEST ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(run_all_tests())
