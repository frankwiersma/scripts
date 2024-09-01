#Scripts

## Email-to-Calendar.ps1 Script

The `Email-to-Calendar.ps1` script automates the process of converting email data into calendar events. To run the script, use the following command in the Windows Run window:

```bash
powershell -ExecutionPolicy Bypass -File "C:\Users\folder\Email-to-Calendar.ps1"
```


## YouTube Subtitles to Clipboard.ps1 Script

The `YouTube Subtitles to Clipboard.ps1` script downloads and processes YouTube subtitles, copying the cleaned text to your clipboard.

### How to Run:

Run the script with this command:

```bash
powershell -ExecutionPolicy Bypass -File "C:\Users\folder\YouTube Subtitles to Clipboard.ps1"
```

### Key Features:

- Downloads English subtitles from a YouTube video using `yt-dlp`.
- Removes timestamps, line numbers, and HTML tags.
- Joins sentences into paragraphs and copies the formatted text to the clipboard.