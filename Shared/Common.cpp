#include "../RedmineDownloader/framework.h"
#include "Common.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


const CString IssueListFileName(_T("\\issue_list.json"));
const CString IssueFileName(_T("\\_issue.json"));
const CString RedmineDataFolder(_T("\\RedmineData"));
const CString MembershipsFileName(_T("\\memberships.json"));
const CString IssueStatusesFileName(_T("\\issue_statuses.json"));
const CString TrackersFileName(_T("\\trackers.json"));
const CString IssuePrioritiesFileName(_T("\\issue_priorities.json"));

const CString IssueListCacheFileName(L"\\issue_Cache.json");

const CString IssueTemplate(_T("/Issue.html"));
const CString LocalHost(_T("https://redmineviewer.example"));

CString GetFileNameFromJson(const CString& fileName, const CString& ID) {
	CString ret;
	ret.Format(L"%s_%s", (LPCWSTR)ID, (LPCWSTR)fileName);
	return ret;
};
