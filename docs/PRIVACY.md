# Student privacy and FERPA

Letter Merger handles student names, scholarships, photos, and personal thank-you messages. Treat all of it, and everything the program creates from it, as confidential. Whether a record is an education record under FERPA depends on 34 CFR 99.3 and how your institution keeps it. The program does not decide that, check anyone's identity, or collect consent.

## What the program does

- Reads CSV files and photos from the local folders and writes the results to a new local batch folder.
- Sends nothing over a network and has no telemetry or accounts.
- Keeps names, scholarships, and complete messages unchanged, apart from the optional encoding repair, which it lists in `TextRepairReport.csv`.
- Redraws photos as new images without their original metadata and removes author and history properties from the merged document.
- Logs only error types and codes at startup, never message text or file paths.

Some of its files are sensitive by design. Reports name students so staff can fix missing photos and layout problems, the missing-photo report covers the whole CSV, and `LetterMergerSettings.json` stores file locations and column names. Word and other apps can keep their own temporary files and recent-file lists.

## What the institution must provide

- **Authority.** Decide who may run merges and confirm the legal basis for sending each letter and photo to a donor. FERPA generally requires signed and dated written consent unless an exception applies, and receiving a scholarship does not by itself allow every donor communication. A form's photo-upload wording may not meet every consent requirement.
- **Devices and storage.** Use a managed device, a restricted account, and encrypted, approved storage. Limit access to the whole working folder, including settings, logs, temporary files, backups, and reports.
- **Review.** Check that each letter goes to the right donor, and read every message for information beyond what the student agreed to share.
- **Retention.** Follow your retention, backup, records-request, secure disposal, and incident response rules. Deleting a file in Windows is not secure erasure, and printed letters need the same care.

Never put real records in GitHub, issue trackers, AI tools, online scanners, or personal cloud accounts unless your institution approves it. A private repository is still an outside service.

These safeguards support careful handling. They are not a FERPA certification or a legal opinion about any institution's use.

## References

Checked October 6, 2026:

- [34 CFR Part 99](https://www.ecfr.gov/current/title-34/subtitle-A/part-99), especially 99.3, 99.10(e), 99.30, 99.31, 99.32, and 99.33
- [U.S. Department of Education, FERPA](https://studentprivacy.ed.gov/ferpa)
- [U.S. Department of Education, Data Security Checklist](https://studentprivacy.ed.gov/resources/data-security-checklist)
