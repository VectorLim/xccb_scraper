import argparse
import json
import logging

from .site import WindowsReader
from .workbook import load_config, report


def main():
    parser = argparse.ArgumentParser(description="Read-only xCCB Flawless Submission reporting")
    parser.add_argument("source", help="Forum Reviewers workbook")
    parser.add_argument("output", help="New report workbook; never overwrites an existing file")
    parser.add_argument("--live", action="store_true", help="Capture fresh Actions before reporting; otherwise use saved evidence")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        counts = report(args.source, args.output, load_config("config.json"), "outputs/evidence",
                        reader=WindowsReader() if args.live else None)
        print(json.dumps(counts))
    except (OSError, ValueError) as exc:
        parser.exit(2, f"Error: {exc}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
