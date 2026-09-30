# Architecture

```text
Synthetic workbook A ----                          > key-indexed comparison -> normalized field checks
Synthetic workbook B ----/                         -> differences
                                                    -> Excel detail report
                                                    -> Excel summary report
```

The comparison logic is deterministic and separated from report generation so both can be tested independently.
