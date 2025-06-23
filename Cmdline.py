#!/usr/bin/env python3

import argparse
from processFiles import processFiles

parser = argparse.ArgumentParser()
parser.add_argument("--defn", required=True, help="Definition filename")
parser.add_argument("--excl", required=True, help="Excel filename")
parser.add_argument("--month", required=True, help="Month")
parser.add_argument("--year", required=True, help="Year")
parser.add_argument("--unit", required=True, help="Business unit")
args = parser.parse_args()

status,result = processFiles(args.defn, args.excl, args.month, args.year, args.unit, True)
print("status = ",status)
print("result = ",result)
