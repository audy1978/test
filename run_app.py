import sys

from streamlit.runtime.runtime import Runtime
from streamlit.web import cli as stcli


if __name__ == "__main__":
    if Runtime.exists():
        import app  # noqa: F401
    else:
        sys.argv = [
            "streamlit",
            "run",
            "app.py",
            "--server.headless=true",
        ]
        sys.exit(stcli.main())
