from pathlib import Path


def test_setup_linux_pins_system_python_for_uv_sync():
    script = Path("desktop/setup_linux.sh").read_text(encoding="utf-8")

    assert "uv venv --python /usr/bin/python3 --no-managed-python --system-site-packages" in script

    sync_lines = [
        line.strip()
        for line in script.splitlines()
        if line.strip().startswith("uv sync")
    ]
    assert sync_lines == [
        "uv sync --python /usr/bin/python3 --no-managed-python --python-preference only-system"
    ]
    assert "uv sync --active" not in script
    assert "include-system-site-packages = true" in script
