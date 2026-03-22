#pragma once

extern const CString IssueListFileName;
extern const CString IssueFileName;
extern const CString RedmineDataFolder;
extern const CString MembershipsFileName;
extern const CString IssueStatusesFileName;
extern const CString TrackersFileName;
extern const CString IssuePrioritiesFileName;

extern const CString IssueListCacheFileName;

extern const CString IssueTemplate;
extern const CString LocalHost;

CString GetFileNameFromJson(const CString& fileName, const CString& ID);
