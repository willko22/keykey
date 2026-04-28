# keykey — Keypress Capture for Usage Statistics

> Minimal, vibecoded keypress collector — counts keys, not chronology.

## Summary

`keykey` is a tiny keypress-capture app that records which keys were pressed for simple usage statistics. It does NOT capture order or timestamps — only that a button was pressed.

This project is intentionally minimal and "vibecoded": small, readable, and focused on the core idea.

## Features

- Records key press occurrences (presence/count) only.
- No timestamps, no ordering — privacy-friendly and lightweight.
- Simple files for config and stats: `config.toml`, `prefs.json`, `stats.json`.

## Files

- `keykey.py` — main script (entrypoint).
- `build_keykey.bat` — convenience Windows launcher.
- `config.toml` — configuration options.
- `prefs.json` — user preferences.
- `stats.json` — output file with key usage counts.

## How it works

When a key is pressed, the app increments that key's count (or marks it pressed). It intentionally does not store the time of the event or the order between events — only that the key was used.

Example `stats.json` (conceptually):

```
{
  "A": 42,
  "Enter": 7,
  "Space": 120
}
```

## Run

On Windows you can run the bundled batch file:

```powershell
.\build_keykey.bat
```

Or run directly with Python:

```bash
python keykey.py
```

Ensure you have a compatible Python installation if running `keykey.py` directly.

## Configuration

Edit `config.toml` or `prefs.json` to adjust behavior (if supported). The app is designed to be plug-and-play; defaults should work out of the box.

## Privacy

This tool purposefully avoids recording timestamps or sequences to reduce sensitive data collection. Use responsibly.

## Contributing

Small, focused patches welcome — keep the vibe minimal.

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License. See [LICENSE](LICENSE) for details.
