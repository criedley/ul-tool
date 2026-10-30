## Time handling (CTS + Agilent)

- `Date/Time` (CTS column 1) is the authoritative absolute timestamp. It is stored
  at full second precision; the source Date/Time has second precision and is in 
  M/D/YYYY h:mm:ss AM/PM format.
  Always read the underlying cell value, never the displayed text.
- `Time` (CTS column 2, units "(sec)") is CTS-provided elapsed seconds. It does not
  start at 0 (first sample is ~1.6 s). Read as-is; do not recompute or zero it.
- Chart x-axis defaults to elapsed seconds (matches the legacy macro, which plots
  column 2). Configurable via `x_unit: sec | min | hr` in config/charts.yaml.
- Agilent has no elapsed column: compute elapsed seconds from cropped timestamps,
  relative to CTS test start.
- Locate columns by header name in row 1 (`Date/Time`, `Time`, `Z_UUT_*_TC_NNN`),
  not by position and not by row-2 descriptions (some TC descriptions are mislabeled).
- First data row = first row where `Date/Time` parses as a datetime. This skips the
  description, units, AVE/MIN/MAX and blank rows. Compute statistics from data
  rows; ignore the file's AVE/MIN/MAX rows.
- Do not depend on unit text (degree symbols arrive as "?").