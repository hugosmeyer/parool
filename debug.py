"""Debug utility module.

This module provides a simple debug print function that can be toggled on/off.
Note: This is deprecated in favor of using the logging module in processFiles.py
"""

import logging

# Configure debug flag
DEBUG_ACTIVE = True

# Set up logger
logging.basicConfig(level=logging.DEBUG if DEBUG_ACTIVE else logging.INFO)
logger = logging.getLogger(__name__)


def debug(*args):
    """Print debug messages if debug mode is active.
    
    Args:
        *args: Variable arguments to print
        
    Note: This function is deprecated. Use the logging module instead.
    """
    if DEBUG_ACTIVE:
        print(*args)
