"""Entry point for the windowed Windows executable."""
import sys

if __name__ == "__main__":
    if len(sys.argv) == 4 and sys.argv[1] == "--organization-worker":
        from xccb_report.organization_worker import main
        raise SystemExit(main())
    else:
        from xccb_report.gui import main
        main()
