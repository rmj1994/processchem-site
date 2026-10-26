#!/usr/bin/env python3
"""Refresh the app screenshots on processchem.no from the app itself.

    python3 tools/screens.py [--app PATH] [--wiki PATH] [--from DIR]

Starts the app's web build (`flet run --web --port 8571 main.py`), drives it with the
app's own `tools/app_shots.py` (headless Chrome, 390 CSS px wide at 2x), then makes:

- `assets/screens/<name>.jpg`: 774 x 1674, the phone frames on the landing page;
- `<wiki>/images/screens/<name>.jpg`: 600 px wide, the pictures in the documentation
  (build_docs.py copies the wiki's images/ into docs/images/).

Each picture is the first phone-height of the screen with its pinned bottom bars (the
Calculate button and/or the tab bar) kept at the foot, as on a phone. `--from DIR` skips
the capture and reuses PNGs already in DIR. Rerun after any change to the app's screens,
then rebuild with build_docs.py. Needs Pillow; the capture also needs Google Chrome,
flet and the app's `websocket-client`.
"""

import argparse
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

from PIL import Image

SITE = Path(__file__).resolve().parent.parent
DEFAULT_APP = Path.home() / "Documents/GitHub/PocketChem-ios"
DEFAULT_WIKI = Path.home() / "Documents/GitHub/ProcessChem-Doc.wiki"
PORT = 8571

PHONE_H = 1688          # 844 CSS px at 2x: one phone screen
LANDING = (774, 1674)   # the size the landing page's <img> tags declare
DOCS_W = 600
BG = (248, 250, 255)

#: app_shots name -> (site name, CSS px of pinned bars at the foot of the capture).
#: Inputs keep the Calculate bar and the tab bar (156); results and lists the tab bar
#: (80); sheets are captured at phone height already (0).
SHOTS = {
    "property_inputs": ("property-inputs", 156),
    "property_results": ("property-results", 80),
    "source_sheet": ("property-source", 0),
    "hazard_sheet": ("property-hazards", 0),
    "menu": ("menu", 0),
    "save": ("save", 0),
    "thermo": ("thermo-modes", 80),
    "case": ("case-list", 80),
    "units": ("units", 80),
    "converter": ("units-converter", 0),
    "phase_inputs": ("phase-inputs", 156),
    "phase_results": ("phase-results", 80),
    "equilibrium_inputs": ("equilibrium-inputs", 156),
    "equilibrium_results": ("equilibrium-results", 80),
    "reaction_inputs": ("reaction-inputs", 156),
    "reaction_results": ("reaction-results", 80),
    "electro_inputs": ("electro-inputs", 156),
    "electro_results": ("electro-results", 80),
    "heater_inputs": ("heater-inputs", 156),
    "heater_results": ("heater-results", 80),
    "vessel_inputs": ("vessel-inputs", 156),
    "vessel_results": ("vessel-results", 80),
    "fluidbed_inputs": ("fluidbed-inputs", 156),
    "fluidbed_results": ("fluidbed-results", 80),
    "hx_inputs": ("hx-inputs", 156),
    "hx_results": ("hx-results", 80),
}

#: The landing page shows these (site names); everything goes to the docs.
LANDING_SHOTS = ("heater-results", "property-results", "thermo-modes", "case-list", "units")


def phone_frame(path: Path, bar_css: int) -> Image.Image:
    """The first phone-height of a capture, its pinned bars moved to the foot."""
    image = Image.open(path).convert("RGB")
    w, h = image.size
    bar = bar_css * 2
    out = Image.new("RGB", (w, PHONE_H), BG)
    out.paste(image.crop((0, 0, w, min(h, PHONE_H) - bar)), (0, 0))
    if bar:
        out.paste(image.crop((0, h - bar, w, h)), (0, PHONE_H - bar))
    return out


def capture(app: Path, out: Path) -> None:
    server = subprocess.Popen(["flet", "run", "--web", "--port", str(PORT), "main.py"], cwd=app,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(60):
            try:
                urllib.request.urlopen(f"http://localhost:{PORT}/", timeout=2)
                break
            except OSError:
                time.sleep(2)
        else:
            sys.exit("the app's web build did not start")
        subprocess.run([sys.executable, str(app / "tools/app_shots.py"), str(out)], cwd=app, check=True)
    finally:
        server.terminate()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--app", type=Path, default=DEFAULT_APP)
    ap.add_argument("--wiki", type=Path, default=DEFAULT_WIKI)
    ap.add_argument("--from", dest="src", type=Path, help="reuse PNGs already captured here")
    args = ap.parse_args()

    src = args.src
    if src is None:
        src = Path(tempfile.mkdtemp(prefix="processchem_shots_"))
        capture(args.app, src)
    missing = [n for n in SHOTS if not (src / f"{n}.png").exists()]
    if missing:
        sys.exit("missing captures: " + ", ".join(missing))

    landing_dir = SITE / "assets" / "screens"
    docs_dir = args.wiki / "images" / "screens"
    landing_dir.mkdir(parents=True, exist_ok=True)
    docs_dir.mkdir(parents=True, exist_ok=True)
    for app_name, (name, bar) in SHOTS.items():
        frame = phone_frame(src / f"{app_name}.png", bar)
        docs = frame.resize((DOCS_W, round(frame.height * DOCS_W / frame.width)), Image.LANCZOS)
        docs.save(docs_dir / f"{name}.jpg", quality=80, optimize=True, progressive=True)
        if name in LANDING_SHOTS:
            frame.resize(LANDING, Image.LANCZOS).save(
                landing_dir / f"{name}.jpg", quality=82, optimize=True, progressive=True)
    print(f"{len(SHOTS)} docs pictures -> {docs_dir}")
    print(f"{len(LANDING_SHOTS)} landing pictures -> {landing_dir}")


if __name__ == "__main__":
    main()
