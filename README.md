# RedmineDownloader & RedmineViewer
RedmineDownloader is a GUI tool designed to batch download and back up ticket information, attachments from Redmine.

RedmineViewer is a GUI tool designed to view & search the downloaded data.

![How it looks](RedmineDownloader.png)

![How it looks](RedmineViewer.png)


## Features - RedmineDownloader
* Simple
  Export tickets with straightforward commands and a user-friendly interface.
* Flexible
  Uses a configuration file located in the same folder as the executable. By placing the executable in different project folders, you can easily manage backups for multiple projects.
* History Tracking
  When an issue is updated, the tool can save a copy under a different name while updating the main issue file. This provides strong protection against "blank page attacks" (accidental or malicious data erasure).
* Designed for Automated Daily Tasks
  Supports command-line options to automate downloads and even shut down the software upon completion.
* Interval
  You can set an interval between access to reduce server load, like 0.5 sec.
* Long Path
  The software can handle filenames and paths exceeding 255 characters.
* Prerequisites
  Requires Redmine API access permissions (API Key).
* License
  MIT License

## Features - RedmineViewer
* Simple
  Just place the executable and RedmineData folder to the folder where you have downloaded the Redmine data, and execute.
* Minimum, but enough functionality
  you can:
  ** show a list of issues
  ** click links to other issues
  ** show the issue; markdown supported
  ** open attachments
  ** search within the Redmine project
  ** can zoom, change dialog size
* Flexible
  HTML template for showing issues, search results and project page can be edited. "Reload" button can reload all templates, and if the active tab is showing an issue, refleshs the page automatically. For other types, you need to add tab to check the edit result.
* Long Path
  The software can handle filenames and paths exceeding 255 characters.
* Speed
  Texts of the issues are cached for search. The cache is updated automatically on boot of the software. If the cache seems to be broken, click Recache button.
* License
  MIT License

## Command Line Options for RedmineDownloader
* -x
  Automatically starts the download process immediately after the software launches. The software will automatically exit once the download is complete.

## RedmineDownloader History
* 1.5.0
** Moved issue_list.json to RedmineData folder
* 1.4.0
** Naming rule of the attahed files are changed. Hence, need to re-download the project.
** List of downloaded items are cached to speed-up existance/update check.
* 1.3.0
** All json files are now saved with UTF-8 without BOM.
** Can select whether to save old data or not.
** Issue.json is removed if failed to get attachments or issues. It will be re-downloaded on next try.
** Downloads memberships.json, statuses.json, trackers.json, and issue_properties.json into folder: RedmineData
* 1.2.0
** Added boot option.
*** -x: Automatically start download, and closes when done.
** Support long path.
* 1.0.0
** Initial Release

## RedmineViewer History
* 2.0.0
** Search is available.
** Cache of all texts supported.
** Re-cache supported.
** Start page is changed to project info.
* 1.2.0
** Changed to tab browser.
* 1.1.0
** Drag & drop supported
** Supports markdown
** WebView2.dll is not required anymore (statically linked)
** Attachment file naming rule is now based on RedmineDownloader 1.4.0.
** Supports opening attachment file.
** Supports jumping to other issues via link.
** Back and forward is now available.
** Maximum / minimum button available.
* 1.0.0
** Initial Release

