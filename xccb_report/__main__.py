import argparse
import json
import logging

from .config import load_config
from .pipeline import report
from .site import WindowsReader


def main():
    parser = argparse.ArgumentParser(
        description="Read-only xCCB Flawless Submission reporting"
    )
    parser.add_argument("source", help="Forum Reviewers workbook")
    parser.add_argument(
        "output", help="New report workbook; never overwrites an existing file"
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="Read live Actions (fresh verified cache is optional); otherwise use saved evidence",
    )
    parser.add_argument("--transport", choices=("session", "powershell"))
    parser.add_argument("--interval", type=float, help="Seconds between request starts (0.1-60)")
    parser.add_argument("--workers", type=int, help="Concurrent document workers (1-16)")
    cache = parser.add_mutually_exclusive_group()
    cache.add_argument("--reuse-evidence", action="store_true", help="Reuse fresh validated evidence")
    cache.add_argument("--force-refresh", action="store_true", help="Bypass the live evidence cache")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        config = load_config("config.json")
        for key, value in (
            ("http_transport", args.transport),
            ("request_interval_seconds", args.interval),
            ("document_workers", args.workers),
        ):
            if value is not None:
                config[key] = value
        if args.reuse_evidence or args.force_refresh:
            config["reuse_evidence"] = args.reuse_evidence
        counts = report(
            args.source,
            args.output,
            config,
            "outputs/evidence",
            reader=WindowsReader(
                interval=config["request_interval_seconds"],
                transport=config["http_transport"],
            ) if args.live else None,
        )
        print(json.dumps(counts))
    except (OSError, ValueError) as exc:
        parser.exit(2, f"Error: {exc}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
