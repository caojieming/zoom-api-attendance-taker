# Zoom Meetings Archiver

Formerly known as the 2 repositories zoom-api-attendance-taker and zoom-api-transcript-getter. I just merged the 2 into 1 for convenience.

This repo is mostly to backup my code, but is public if anyone else needs this oddly specific usecase (I've scrubbed all personal IDs/API keys).

## Zoom API Attendance Taker

Google Apps Script to GET past meetings + participants from Zoom API and write all that info into a Google Sheet, with each meeting occupying a distinct sheet tab. Comes with an additional function that generates a tab with attendance rankings by attendance rate.

## Zoom API Recordings/Transcripts Getter

Google Apps Script to GET meeting transcripts from Zoom API and write them into a Google Doc in Drive. Also generates a sheet for each meeting that contains meeting topics.