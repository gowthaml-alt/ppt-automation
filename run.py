"""Entry point used by Windows Task Scheduler / NSSM."""

from main import main


if __name__ == "__main__":
    raise SystemExit(main())
