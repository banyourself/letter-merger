# Security

## Reporting a problem

Email me at kevin@kevinle.tech with a description, the version, and steps to reproduce it using made-up data. Please do not open a public issue for a security problem, and never attach real student records, photos, letters, or screenshots that show them.

If real student records were exposed, also follow your institution's incident response procedure. The program cannot report incidents or decide what may be disclosed.

## How the program protects data

- **Local only.** There is no network code, telemetry, account, credential storage, or email sending. Opening an output uses the file's registered Windows app.
- **CSV input.** Files over 64 MB and Excel workbooks renamed to `.csv` are refused. The parser is strict about quotes and column counts, and the source CSV is never changed.
- **Photo paths.** Each path must be relative and stay inside `Photos`. Absolute paths, drive letters, URLs, `..` segments, control characters, and file system links are refused.
- **Photos.** Files over 64 MB or 64 million pixels are refused, and the file signature must match a supported raster format, so a renamed vector or metafile is rejected. Photos are redrawn as new PNG images, which drops their original metadata.
- **Template.** Only the approved blank card is accepted. The program checks its part list, relationships, merge fields, and text, pins the frame image by hash, and refuses macros, embedded objects, external links, active Word fields, tracked changes, comments, and hidden text. XML parsing has DTDs and external entities turned off, and DOCX files have part-count and size limits.
- **Output.** Each run writes to a new batch folder, and merged documents and PDFs never overwrite an existing file. Author and history properties are replaced with minimal metadata. Every card is checked so its name, scholarship, and complete message match the CSV.
- **Word.** Word opens the merged file read-only, without adding it to recent files, with macros forced off.
- **Reports.** Every cell is quoted, and values that start with `=`, `+`, `-`, or `@` are prefixed so a spreadsheet shows them as text.
- **Logs.** The startup log records only the error type and code, never message text or file paths.

## Limits

- The program is not a sandbox for hostile files. It relies on Windows, .NET Framework, GDI+, Word, and any installed image codecs, so keep them updated.
- It runs with the current user's permissions and does not encrypt files or control who can open them. Use a managed device and a folder only authorized staff can reach.
- Interrupted runs can leave partial files in the batch folder, and Word can keep its own temporary files.
- There is no time limit on Word automation, so a stuck Word window can stall the layout check.
- The executable is not code signed.

## Before publishing a change

Run `scripts/AuditRepository.py` in all three modes (see the README) and look over the staged file list. The `.gitignore` file allows only the approved files, but a forced add can bypass it, and nothing removes a file from Git history after it is committed. Pattern checks cannot find every kind of personal information, so a person still needs to review each change.
