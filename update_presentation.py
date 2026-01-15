#!/usr/bin/env python3
"""
Google Sheets -> PowerPoint Investor Deck Auto-Update
Main orchestrator script that pulls data from Google Sheets and updates PowerPoint.

Usage:
    python update_presentation.py [--config CONFIG] [--dry-run] [--verbose]
"""

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import Optional

import yaml

from sheets_client import SheetsClient, MappingRow, create_client
from format_value import format_value
from powerpoint_bridge import PowerPointBridge, check_presentation

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).parent


def load_config(config_path: str) -> dict:
    """Load configuration from YAML file."""
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"Configuration file '{config_path}' not found.")

    with open(config_file, 'r') as f:
        config = yaml.safe_load(f)

    logger.debug(f"Loaded configuration from {config_path}")
    return config


def process_mapping(bridge: PowerPointBridge, mapping: MappingRow, value, config: dict,
                    dry_run: bool = False) -> tuple[bool, str]:
    """
    Process a single mapping: format the value and update PowerPoint.
    """
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


def main():
    """Main entry point for the update script."""
    parser = argparse.ArgumentParser(
        description='Update PowerPoint deck from Google Sheets data'
    )
    parser.add_argument(
        '--config', '-c',
        default='config.yaml',
        help='Path to configuration file (default: config.yaml)'
    )
    parser.add_argument(
        '--dry-run', '-n',
        action='store_true',
        help='Show what would be updated without making changes'
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose/debug logging'
    )
    parser.add_argument(
        '--presentation', '-p',
        help='Path to PowerPoint file (overrides config)'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output path for updated presentation (default: overwrite original)'
    )
    parser.add_argument(
        '--list-shapes',
        type=int,
        metavar='SLIDE',
        help='List named shapes on specified slide (for setup)'
    )
    parser.add_argument(
        '--list-tables',
        type=int,
        metavar='SLIDE',
        help='List tables on specified slide (for setup)'
    )

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    os.chdir(SCRIPT_DIR)

    # Load configuration
    try:
        config = load_config(args.config)
    except Exception as e:
        logger.error(f"Failed to load configuration: {e}")
        sys.exit(1)

    # Get presentation path
    pptx_path = args.presentation or config.get('powerpoint', {}).get('file_path', '')
    if not pptx_path:
        logger.error("No PowerPoint file specified. Use --presentation or set powerpoint.file_path in config.")
        sys.exit(1)

    # Handle utility commands
    if args.list_shapes or args.list_tables:
        bridge = PowerPointBridge(pptx_path)
        if not bridge.open():
            logger.error(f"Failed to open PowerPoint: {pptx_path}")
            sys.exit(1)

        if args.list_shapes:
            shapes = bridge.list_shapes(args.list_shapes)
            print(f"\nShapes on slide {args.list_shapes}:")
            for s in shapes:
                print(f"  - {s['name']} ({s['type']}) {'[text]' if s['has_text'] else ''} {'[table]' if s['has_table'] else ''}")
                if s.get('text_preview'):
                    print(f"      Text: {s['text_preview']}...")

        if args.list_tables:
            tables = bridge.list_tables(args.list_tables)
            print(f"\nTables on slide {args.list_tables}:")
            for t in tables:
                print(f"  - {t['name']} ({t['rows']} rows x {t['columns']} cols)")

        sys.exit(0)

    # Check presentation exists
    if not Path(pptx_path).exists():
        logger.error(f"PowerPoint file not found: {pptx_path}")
        sys.exit(1)

    dry_run = args.dry_run or config.get('dry_run', False)
    if dry_run:
        logger.info("Running in DRY RUN mode - no changes will be made")

    # Open presentation
    bridge = PowerPointBridge(pptx_path)
    if not bridge.open():
        logger.error(f"Failed to open PowerPoint: {pptx_path}")
        sys.exit(1)

    logger.info(f"Opened presentation: {pptx_path} ({bridge.get_slide_count()} slides)")

    # Initialize Sheets client
    try:
        sheets_client = create_client(config)
        logger.info("Initialized Google Sheets client")
    except Exception as e:
        logger.error(f"Failed to initialize Sheets client: {e}")
        sys.exit(1)

    # Read mapping configuration
    try:
        mappings = sheets_client.read_mapping()
        if not mappings:
            logger.warning("No mappings found in KeynoteMap sheet")
            sys.exit(0)
        logger.info(f"Found {len(mappings)} mappings to process")
    except Exception as e:
        logger.error(f"Failed to read mappings: {e}")
        sys.exit(1)

    # Collect unique ranges to fetch
    ranges = sorted(set(m.sheet_range for m in mappings if m.sheet_range))
    if not ranges:
        logger.warning("No valid sheet ranges found in mappings")
        sys.exit(0)

    # Batch fetch all values from Google Sheets
    try:
        values_by_range = sheets_client.batch_get_values(ranges)
        logger.info(f"Fetched {len(values_by_range)} values from Google Sheets")
    except Exception as e:
        logger.error(f"Failed to fetch values from Sheets: {e}")
        sys.exit(1)

    # Process each mapping
    success_count = 0
    error_count = 0
    errors = []

    for mapping in mappings:
        raw_value = values_by_range.get(mapping.sheet_range)
        success, message = process_mapping(bridge, mapping, raw_value, config, dry_run)

        if success:
            success_count += 1
            logger.info(f"Updated '{mapping.id}': {message}")
        else:
            error_count += 1
            error_msg = f"Failed to update '{mapping.id}': {message}"
            errors.append(error_msg)
            logger.error(error_msg)

    # Save presentation
    if not dry_run and success_count > 0:
        output_path = args.output or pptx_path
        if bridge.save(output_path):
            logger.info(f"Saved presentation to: {output_path}")
        else:
            logger.error("Failed to save presentation")
            error_count += 1

    # Summary
    print()
    print("=" * 50)
    print(f"Update Complete{' (DRY RUN)' if dry_run else ''}")
    print(f"  Successful: {success_count}")
    print(f"  Failed: {error_count}")
    print("=" * 50)

    if errors:
        print("\nErrors:")
        for err in errors:
            print(f"  - {err}")

    sys.exit(0 if error_count == 0 else 1)


if __name__ == '__main__':
    main()
