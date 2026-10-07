# Call_Log_5
Homeroom Automations v5

## Grade report photos

Add `GradeReportPhotos.gs` to the **bound Apps Script project for 10RosterHR**
and replace `Menus.gs` with this version. Keep `Config.gs` and the other files.
Reload the spreadsheet to see the **Grade Reports** menu.

Choose **Import grade report photos**, paste the week's Drive folder link, and
enter **Z** for Week 7 (the existing header is `Week 7 of 10/5`). Approve Google's
Drive/Sheets authorization when first requested. The script needs read access to
the folder and edit access to the spreadsheet. GitHub updates alone do not deploy
code into the bound Apps Script project.

- When Grades has only headers, it copies student details from Master into A:F.
  It retains existing rows and course data on subsequent imports.
- It matches the filename's name against each student's preferred and legal
  names, including multipart names. The Canvas number before `_question_` is
  **not** treated as the roster Student Number. Ambiguous names, duplicate photos,
  unmatched files, and unsupported types are reported/skipped.
- PNG, JPEG, and GIF files are embedded directly from private Drive blobs as
  **images over cells**, anchored in the weekly column, with aspect ratio retained.
  The photo column is 420 pixels wide; matched rows are at least 240 pixels high.
  Existing cell text, hyperlinks, notes, and formulas are retained. Original-photo
  links appear in the import/preview report.
- If the original exceeds the byte or pixel insertion limits, it automatically
  tries Drive's private thumbnail blob. Originals over 2 MB go straight to the
  thumbnail; smaller files retry with a thumbnail only on a size/pixel-limit error.
  The report identifies thumbnail previews; open the original report for readable
  full-resolution text. A missing or unusable thumbnail is reported; previous
  photos and original Drive files remain intact. Preview mode checks name matches
  only, not whether Google's image service will accept each image.
- If both original and thumbnail exceed the pixel limit, the import report
  automatically resizes the image in your browser to a maximum of 900 pixels on
  either side (at most 810,000 pixels). PNG output is reduced further if needed
  to stay below 2 MB. **Keep the report dialog open until resizing finishes.**
  Server-side checks verify the actual PNG dimensions, byte size, pending student,
  current roster row, weekly header, and original photo revision before insertion.
  It needs no new service, third-party upload, or sharing change. Each resized
  insertion gets a `GradeReportPhotosResize` success entry in Executions. Already
  imported photos are retained. The original-file links still open full-resolution
  reports; the initial WARN entry may precede successful resize callback entries.
- Repeating an import does not duplicate the script's images. New files replace
  only images this tool previously created for that student in that column.
  Missing photos do not clear prior photos. Manually inserted images are retained.
- Use **Preview grade report photos** to see matches without modifying the sheet.
  This is a menu-driven import, not a scheduled task. Next week, paste the new
  folder and choose its weekly column. Folder IDs remain in document properties;
  no student data or private folder links belong in this public repository.

Run the local matching/import-behavior tests with `node --test tests/grade-report-photos.test.cjs`.
Optional native-Canvas integration test (requires `@napi-rs/canvas`):
`node tests/canvas-grade-photo-resize.cjs [photo.png ...]`. It executes the dialog's
actual resize/save JavaScript with PNG input files, without calling live Sheets.
