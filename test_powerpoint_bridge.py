#!/usr/bin/env python3
"""
Test script for PowerPoint Bridge module.
Tests reading and updating shapes and tables in PowerPoint files.
"""

import logging
import sys
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Add project directory to path
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

from powerpoint_bridge import PowerPointBridge, check_presentation


def test_open_presentation():
    """Test opening a PowerPoint file."""
    print("\n" + "=" * 60)
    print("TEST: Opening presentation")
    print("=" * 60)

    pptx_path = script_dir / "sample_investor_deck.pptx"

    # Test check_presentation function
    success, message = check_presentation(str(pptx_path))
    print(f"check_presentation: {success} - {message}")
    assert success, f"Failed to check presentation: {message}"

    # Test opening with PowerPointBridge
    bridge = PowerPointBridge(str(pptx_path))
    result = bridge.open()
    print(f"bridge.open(): {result}")
    assert result, "Failed to open presentation"

    slide_count = bridge.get_slide_count()
    print(f"Slide count: {slide_count}")
    assert slide_count == 5, f"Expected 5 slides, got {slide_count}"

    print("PASS: Opening presentation")
    return bridge


def test_list_shapes(bridge: PowerPointBridge):
    """Test listing shapes on slides."""
    print("\n" + "=" * 60)
    print("TEST: Listing shapes")
    print("=" * 60)

    # Test slide 1
    shapes = bridge.list_shapes(1)
    print(f"\nSlide 1 shapes ({len(shapes)}):")
    for s in shapes:
        print(f"  - {s['name']}: {s['type']}, text={s['has_text']}")

    assert len(shapes) >= 2, "Expected at least 2 shapes on slide 1"
    shape_names = [s['name'] for s in shapes]
    assert 'Title' in shape_names, "Expected 'Title' shape on slide 1"
    assert 'ReportDate' in shape_names, "Expected 'ReportDate' shape on slide 1"

    # Test slide 2
    shapes = bridge.list_shapes(2)
    print(f"\nSlide 2 shapes ({len(shapes)}):")
    for s in shapes:
        print(f"  - {s['name']}: {s['type']}, text={s['has_text']}")

    shape_names = [s['name'] for s in shapes]
    assert 'RevenueValue' in shape_names, "Expected 'RevenueValue' shape on slide 2"
    assert 'GrowthValue' in shape_names, "Expected 'GrowthValue' shape on slide 2"

    print("\nPASS: Listing shapes")


def test_list_tables(bridge: PowerPointBridge):
    """Test listing tables on slides."""
    print("\n" + "=" * 60)
    print("TEST: Listing tables")
    print("=" * 60)

    # Test slide 3 (has FinancialTable)
    tables = bridge.list_tables(3)
    print(f"\nSlide 3 tables ({len(tables)}):")
    for t in tables:
        print(f"  - {t['name']}: {t['rows']}x{t['columns']}")

    assert len(tables) >= 1, "Expected at least 1 table on slide 3"
    table_names = [t['name'] for t in tables]
    assert 'FinancialTable' in table_names, "Expected 'FinancialTable' on slide 3"

    # Test slide 4 (has KPITable)
    tables = bridge.list_tables(4)
    print(f"\nSlide 4 tables ({len(tables)}):")
    for t in tables:
        print(f"  - {t['name']}: {t['rows']}x{t['columns']}")

    assert len(tables) >= 1, "Expected at least 1 table on slide 4"
    table_names = [t['name'] for t in tables]
    assert 'KPITable' in table_names, "Expected 'KPITable' on slide 4"

    print("\nPASS: Listing tables")


def test_update_shape_text(bridge: PowerPointBridge):
    """Test updating shape text."""
    print("\n" + "=" * 60)
    print("TEST: Updating shape text")
    print("=" * 60)

    # Test updating RevenueValue on slide 2
    success, message = bridge.update_shape_text(2, "RevenueValue", "$6,500,000")
    print(f"Update RevenueValue: {success} - {message}")
    assert success, f"Failed to update RevenueValue: {message}"

    # Test updating GrowthValue on slide 2
    success, message = bridge.update_shape_text(2, "GrowthValue", "32.5%")
    print(f"Update GrowthValue: {success} - {message}")
    assert success, f"Failed to update GrowthValue: {message}"

    # Test updating Title on slide 1
    success, message = bridge.update_shape_text(1, "Title", "Updated Investor Report")
    print(f"Update Title: {success} - {message}")
    assert success, f"Failed to update Title: {message}"

    # Test updating ReportDate on slide 1
    success, message = bridge.update_shape_text(1, "ReportDate", "Q1 2025")
    print(f"Update ReportDate: {success} - {message}")
    assert success, f"Failed to update ReportDate: {message}"

    # Test updating non-existent shape
    success, message = bridge.update_shape_text(1, "NonExistentShape", "Test")
    print(f"Update NonExistentShape (expected fail): {success} - {message}")
    assert not success, "Expected failure for non-existent shape"

    # Test updating shape on non-existent slide
    success, message = bridge.update_shape_text(99, "Title", "Test")
    print(f"Update on slide 99 (expected fail): {success} - {message}")
    assert not success, "Expected failure for non-existent slide"

    print("\nPASS: Updating shape text")


def test_update_table_cell(bridge: PowerPointBridge):
    """Test updating table cells."""
    print("\n" + "=" * 60)
    print("TEST: Updating table cells")
    print("=" * 60)

    # Test updating FinancialTable on slide 3
    # Cell (2, 4) = Q4 2024 Revenue = "$5,000,000"
    success, message = bridge.update_table_cell(3, "FinancialTable", 2, 4, "$7,500,000")
    print(f"Update FinancialTable (2,4): {success} - {message}")
    assert success, f"Failed to update table cell: {message}"

    # Test updating KPITable on slide 4
    # Cell (2, 2) = Current CAC = "$150"
    success, message = bridge.update_table_cell(4, "KPITable", 2, 2, "$120")
    print(f"Update KPITable (2,2): {success} - {message}")
    assert success, f"Failed to update table cell: {message}"

    # Test updating non-existent table
    success, message = bridge.update_table_cell(3, "NonExistentTable", 1, 1, "Test")
    print(f"Update NonExistentTable (expected fail): {success} - {message}")
    assert not success, "Expected failure for non-existent table"

    # Test updating out-of-range row
    success, message = bridge.update_table_cell(3, "FinancialTable", 99, 1, "Test")
    print(f"Update out-of-range row (expected fail): {success} - {message}")
    assert not success, "Expected failure for out-of-range row"

    # Test updating out-of-range column
    success, message = bridge.update_table_cell(3, "FinancialTable", 1, 99, "Test")
    print(f"Update out-of-range col (expected fail): {success} - {message}")
    assert not success, "Expected failure for out-of-range column"

    print("\nPASS: Updating table cells")


def test_save_presentation(bridge: PowerPointBridge):
    """Test saving the presentation."""
    print("\n" + "=" * 60)
    print("TEST: Saving presentation")
    print("=" * 60)

    output_path = script_dir / "test_output.pptx"
    success = bridge.save(str(output_path))
    print(f"Save to {output_path}: {success}")
    assert success, "Failed to save presentation"

    # Verify the file exists
    assert output_path.exists(), "Output file does not exist"
    print(f"Output file size: {output_path.stat().st_size} bytes")

    # Re-open and verify changes were saved
    bridge2 = PowerPointBridge(str(output_path))
    assert bridge2.open(), "Failed to re-open saved presentation"

    # Check slide count
    assert bridge2.get_slide_count() == 5, "Slide count mismatch after save"

    print("\nPASS: Saving presentation")

    # Clean up
    output_path.unlink()
    print(f"Cleaned up: {output_path}")


def run_all_tests():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("POWERPOINT BRIDGE MODULE TESTS")
    print("=" * 60)

    try:
        bridge = test_open_presentation()
        test_list_shapes(bridge)
        test_list_tables(bridge)
        test_update_shape_text(bridge)
        test_update_table_cell(bridge)
        test_save_presentation(bridge)

        print("\n" + "=" * 60)
        print("ALL TESTS PASSED!")
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
