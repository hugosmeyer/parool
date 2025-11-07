#!/usr/bin/env python3
"""Command-line interface for payroll file processing.

This script provides a command-line interface to process Excel payroll files
according to definition file specifications.
"""

import argparse
import sys
from processFiles import processFiles


def main():
    """Parse command-line arguments and process files."""
    parser = argparse.ArgumentParser(
        description="Process Excel payroll files according to definition specifications"
    )
    parser.add_argument("--defn", required=True, help="Definition filename (INI format)")
    parser.add_argument("--excl", required=True, help="Excel filename (.xlsx)")
    parser.add_argument("--month", required=True, help="Month (e.g., Jan, Feb, Mar)")
    parser.add_argument("--year", required=True, help="Year (e.g., 2025)")
    parser.add_argument("--unit", required=True, help="Business unit name")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    
    args = parser.parse_args()

    status, result = processFiles(
        args.defn, args.excl, args.month, args.year, args.unit, args.debug
    )
    
    print(f"Status: {status}")
    if result:
        print(f"Result: {result}")
        
    # Exit with error code if processing failed
    if status == "Failed":
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
