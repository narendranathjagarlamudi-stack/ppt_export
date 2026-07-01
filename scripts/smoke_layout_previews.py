import json
import sys
import urllib.error
import urllib.request


BASE_URL = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8000"
URL = f"{BASE_URL.rstrip('/')}/export/layouts?refresh_cache=true"


def main():
    try:
        with urllib.request.urlopen(URL, timeout=180) as response:
            payload = json.loads(response.read().decode("utf-8"))

    except urllib.error.HTTPError as exc:
        print(f"HTTP {exc.code}: {exc.read().decode('utf-8', errors='ignore')}")
        return 1

    except Exception as exc:
        print(f"Request failed: {type(exc).__name__}: {exc}")
        return 1

    layouts = payload.get("layouts") or []
    mode = payload.get("preview_mode")
    renderer = payload.get("renderer")
    error = payload.get("renderer_error")

    print(f"preview_mode={mode}")
    print(f"renderer={renderer}")
    print(f"layout_count={len(layouts)}")

    if error:
        print(f"renderer_error={error}")

    if mode != "libreoffice":
        print("Expected LibreOffice-rendered thumbnails.")
        return 1

    if not layouts:
        print("No layouts returned.")
        return 1

    first_preview = layouts[0].get("preview", "")

    if not first_preview.startswith("data:image/png;base64,"):
        print("First layout preview is not a PNG data URL.")
        return 1

    print("LibreOffice layout preview smoke test passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
