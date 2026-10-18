# Tools

## `video/` — B1/B2 tutorial video (MP4, ~20 min)

A small engine (`engine.py`) draws a spreadsheet mock-up, slides and subtitles, syncs them to narration from Piper TTS (offline), and pipes frames to ffmpeg (bundled via `imageio-ffmpeg`). `script.py` holds the narration and on-screen actions; formulas and values are read from `../../workbooks/Group_8_local_snapshot_2026-10-11.xlsx`.

```bash
cd tools/video
python -m venv venv && venv/bin/pip install -r requirements.txt
mkdir -p voices && cd voices && for f in en_US-lessac-medium.onnx en_US-lessac-medium.onnx.json; do
  curl -sSLO "https://huggingface.co/rhasspy/piper-voices/resolve/main/en/en_US/lessac/medium/$f"; done && cd ..
venv/bin/python charts.py                 # charts/*.png
venv/bin/python script.py --preview 150 503 835   # still frames -> preview/
venv/bin/python script.py                 # full render -> out/B1_B2_guide.mp4 (~20 min on 4 cores)
```

Fonts used: Ubuntu, Carlito, DejaVu Sans (+ Mono). On Debian/Ubuntu: `fonts-ubuntu fonts-crosextra-carlito fonts-dejavu`.

## `report/` — B1/B2 report sections (.docx)

`export_cells.py` exports the B1/B2 cells (formula + value) from the workbook snapshot to `cells.json`; the docx-js builders read only that file, so every number in the documents matches the workbook.

```bash
cd tools/report
npm install
python export_cells.py
node build_report.js Group_8_B1_B2_report.docx      # detailed, ~6 pages
node build_short.js vi Group_8_B1_B2_short_VI.docx  # short, Vietnamese (decimal commas)
node build_short.js en Group_8_B1_B2_short_EN.docx  # short, English
```
