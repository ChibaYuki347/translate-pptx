## Instruction

Translate Japanese PPTX files in the `pptx/` folder (`*.pptx`, excluding `*_EN.pptx`) into English.

### Steps

1. Enumerate target PPTX files in the `pptx/` folder.
   - Target: `pptx/*.pptx`
   - Excluded: filenames ending with `_EN.pptx` (already-translated English output of a previous ja2en run)
   - `_JA.pptx` files (Japanese output of a prior en2ja run) **are valid inputs** for ja2en and must be included.
2. For each file, **process them one at a time (no parallelism)** and invoke the
   `translate-pptx` agent as a sub-agent.
   - Arguments:
     - **Absolute path** to the source PPTX file
     - `direction`: `ja2en`
     - Output suffix: `_EN`
3. After all files have been processed, report:
   - The list of output files
   - Per-file success / failure (with an error summary on failure)
   - The list of files that were skipped (with reasons)

### Constraints

- **File-level parallelism is forbidden** (process one file at a time, serially).
  - Reason: the `translate-pptx` agent shares the `temp/` folder under the repository root as its working area, so concurrent runs will overwrite or delete each other's intermediate files.
- Do not overwrite the source file. Save the translation as a new file with the `_EN` suffix.
- If one file fails, continue processing the remaining files and report all results at the end.
- The agent rewrites `lang="ja-*"` to `lang="en-US"`, removes East-Asian fonts (`<a:ea>`), and forces the Latin font to **Segoe UI** (body) / **Segoe UI Semibold** (bold / heading) so that titles do not fall back to PowerPoint's default Calibri Light.
