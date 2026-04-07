from pathlib import Path
import sys

from streamlit.web import cli as stcli


def main() -> int:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    launcher_path = base_path / "Launcher.py"
    sys.argv = [
        "streamlit",
        "run",
        str(launcher_path),
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
    ]
    return stcli.main()


if __name__ == "__main__":
    raise SystemExit(main())
