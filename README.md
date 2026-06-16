# Auto Cutdown Tool

A desktop application for batch-cutting video clips based on timecodes defined in a CSV or Excel spreadsheet. Built with Python, Tkinter, and FFmpeg.

## Features

- **Batch Processing** — Cut multiple clips from a single source video in one run.
- **CSV & Excel Support** — Read timecodes from `.csv` or `.xlsx` files.
- **Flexible Timecode Formats** — Supports `HH:MM:SS`, `MM:SS`, raw seconds, and frame-based `HH:MM:SS:FF` (25 fps) formats.
- **Fast Stream Copy** — Uses FFmpeg stream copy (`-c copy`) for near-instant, lossless cutting.
- **Automatic Transcoding Fallback** — If stream copy fails (e.g., ProRes `.mov` to `.mp4`), automatically retries with H.264/AAC transcoding.
- **GUI Interface** — Simple Tkinter-based UI with progress bar and log output.
- **Configurable Output Suffix** — Customize the naming suffix for output clips.

## Requirements

- **Python** 3.8+
- **FFmpeg** — `ffmpeg.exe` must be placed in the same directory as the script (or the built executable).

### Python Dependencies

```
moviepy==1.0.3
pandas
openpyxl
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Running from Source

```bash
python auto_cutdown_tool.py
```

### Building a Standalone Executable

```bash
pyinstaller auto_cutdown_tool.spec
```

The executable will be generated in the `dist/auto_cutdown_tool/` directory. Make sure to copy `ffmpeg.exe` into the same folder as the executable.

### Steps

1. **Select Input Sheet** — Choose a CSV or Excel file containing timecodes.
2. **Select Source Video** — Choose the video file to cut from (supports `.mp4`, `.mov`, `.mxf`, `.avi`, `.mkv`, `.wmv`).
3. **Select Output Folder** — Choose where the cut clips will be saved.
4. **Set Output Suffix** (optional) — Customize the suffix appended to output filenames (default: `_cut`).
5. **Click "Start Processing"** — The tool will process each row and save the clips.

## Input Sheet Format

The input CSV or Excel file must contain columns for start and end timecodes. The following column name pairs are supported (case-insensitive):

| Column Pair          | Start Column   | End Column     |
|----------------------|----------------|----------------|
| Source format        | `source_in`    | `source_out`   |
| Source format (alt)  | `Source In`    | `Source Out`    |
| Timeline format      | `Timeline In`  | `Timeline Out` |

### Example CSV

```csv
No.,source_in,source_out
1,0:00:00,0:00:05
2,0:00:05,0:00:10
3,0:00:10,0:00:15
4,0:00:15,0:00:19
```

### Supported Timecode Formats

| Format          | Example        | Description                     |
|-----------------|----------------|---------------------------------|
| `HH:MM:SS`      | `0:01:30`      | Hours, minutes, seconds         |
| `MM:SS`          | `1:30`         | Minutes, seconds                |
| `HH:MM:SS:FF`   | `00:01:30:12`  | With frames (25 fps)            |
| Seconds (float)  | `90.5`         | Raw seconds                     |

## Output

- Output clips are saved as `.mp4` files with the naming pattern: `{source_video_name}{suffix}{index}.mp4`
- Example: `my_video_cut1.mp4`, `my_video_cut2.mp4`, ...

## Project Structure

```
├── auto_cutdown_tool.py    # Main application (GUI + FFmpeg-based processing)
├── auto_cutdown_tool.spec  # PyInstaller build spec
├── video_cutter.py         # CLI-based video cutter (alternative, uses moviepy)
├── videocutapp.py          # GUI video cutter (alternative, uses moviepy)
├── requirements.txt        # Python dependencies
├── data.csv                # Example input sheet
├── ffmpeg.exe              # FFmpeg binary (not tracked in git)
└── ffprobe.exe             # FFprobe binary (not tracked in git)
```

## Notes

- `ffmpeg.exe` and `ffprobe.exe` are **not included** in this repository due to their large file size. Download them from [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html) and place them in the project directory.
- The tool prioritizes stream copy for speed. If the source codec is incompatible with the MP4 container (e.g., Apple ProRes), it will automatically fall back to transcoding with `libx264` + `aac`.

## License

This project is for internal use.
