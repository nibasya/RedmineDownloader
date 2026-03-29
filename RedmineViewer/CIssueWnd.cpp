#include "pch.h"
#include "CIssueWnd.h"
#include "RedmineViewer.h"
#include "RedmineViewerDlg.h"
#include "RPTT.h"
#include "../Shared/Common.h"
#include <pathcch.h>
#include <regex>

using namespace Microsoft::WRL;
using namespace wil;

#pragma comment(lib, "pathcch")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

BEGIN_MESSAGE_MAP(CIssueWnd, CWnd)
//	ON_WM_DESTROY()
ON_WM_SIZE()
ON_WM_DESTROY()
ON_WM_SHOWWINDOW()
ON_WM_CLOSE()
END_MESSAGE_MAP()

CIssueWnd::CIssueWnd(CRedmineViewerDlg* pParent)
	: m_pParent(pParent), m_WebView(nullptr), m_WebViewController(nullptr)
{
}


CIssueWnd::~CIssueWnd()
{
	if (m_WebViewController) {
		m_WebViewController->Close();
	}
}

bool CIssueWnd::CreateWebView2(UINT nID, LPCWSTR url)
{
	HRESULT hr;
	RECT rect;
	m_pParent->m_CtrlTabViewer.GetClientRect(&rect);
	//if (!Create(NULL, NULL, WS_CHILD | WS_VISIBLE, rect, &m_pParent->m_CtrlTabViewer, nID))
	if (!Create(NULL, NULL, WS_CHILD, rect, &m_pParent->m_CtrlTabViewer, nID))
		return FALSE;

	m_RequestedUrl = url;

	// --- WebView2 初期化（WIL + WRL::Callback を利用） ---
	// Create controller and set up WebView2
	hr = m_pParent->m_pWebEnvironment->CreateCoreWebView2Controller(m_pParent->m_CtrlTabViewer.m_hWnd,
		Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
			[this](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
				return this->WebView2CreateController(result, controller);
			}).Get());
	if (!SUCCEEDED(hr)) {
		UNREFERENCED_PARAMETER(hr);
		ShowMessageBox(L"Failed to create WebView2Control", L"WebView2Error");
		return false;
	}
	return true;
}

HRESULT CIssueWnd::WebView2CreateController(HRESULT result, ICoreWebView2Controller* controller)
{
	if (FAILED(result) || controller == nullptr) return result;

	m_WebViewController = controller;
	ICoreWebView2* rawWebView = nullptr;
	if (SUCCEEDED(m_WebViewController->get_CoreWebView2(&rawWebView)) && rawWebView != nullptr) {
		m_WebView = rawWebView;

		// configure bounds of the WebView2 control to fit the dialog's client area
		CRect rc;
		m_pParent->m_CtrlTabViewer.GetWindowRect(&rc);
		m_pParent->ScreenToClient(&rc);
		RECT bounds = { rc.left, rc.top, rc.right, rc.bottom };
		m_WebViewController->put_Bounds(bounds);

		// prevent new window and get uri when file is drag&dropped
		m_WebView->add_NewWindowRequested(
			Callback<ICoreWebView2NewWindowRequestedEventHandler>(
				[this](ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args) {
					return this->NewWindowRequestHandler(sender, args);
				}).Get(), nullptr);

		// Open local files with default app when a link is clicked
		m_WebView->add_NavigationStarting(
			Callback<ICoreWebView2NavigationStartingEventHandler>(
				[this](ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args) -> HRESULT {
					return this->NavigationStartingHandler(sender, args);
				}).Get(), nullptr);
		// Update button states
		m_WebView->add_NavigationCompleted(
			Callback<ICoreWebView2NavigationCompletedEventHandler>(
				[this](ICoreWebView2* webView, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
					BOOL flag = FALSE;
					webView->get_CanGoBack(&flag);
					m_pParent->m_CtrlButtonBack.EnableWindow(flag ? TRUE : FALSE);
					webView->get_CanGoForward(&flag);
					m_pParent->m_CtrlButtonForward.EnableWindow(flag ? TRUE : FALSE);
					return S_OK;
				}).Get(), nullptr);
		// Update tab title
		m_WebView->add_DocumentTitleChanged(
			Callback<ICoreWebView2DocumentTitleChangedEventHandler>(
				[this](ICoreWebView2* sender, IUnknown* args) -> HRESULT {
					wil::unique_cotaskmem_string title;
					sender->get_DocumentTitle(&title);
					if (title.get()) {
						CString str = title.get();
						m_pParent->m_CtrlTabViewer.SetTabLabel(m_pParent->m_CtrlTabViewer.GetActiveTab(), str);
					}
					return S_OK;
				}).Get(), nullptr);
		CRect rect;
		m_pParent->m_CtrlTabViewer.GetWndArea(rect);
		m_WebViewController->put_Bounds(rect);
		SetWindowPos(&wndBottom, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		CString docType = L"<!doctype html>\n";
		if (m_RequestedUrl.Left(docType.GetLength()) == docType) {
			m_WebView->NavigateToString(m_RequestedUrl);
		}
		else {
			m_WebView->Navigate(m_RequestedUrl);
		}
	}
	return S_OK;
}

bool CIssueWnd::WebView2GetLocalFilePathFromUri(const wil::unique_cotaskmem_string& uri, CString& outPath)
{
	// check if the URI is local file path
	CString path(uri.get());
	if (path.Left(8).CompareNoCase(_T("file:///")) == 0) {
		path = path.Mid(8);	// remove "file:///"
	}
	else {
		ShowMessageBox(CString(L"Only opening local file is allowed: ") + path, L"Error");
		return false;
	}
	const int maxPath = 32767 + 1;
	HRESULT hr = UrlUnescapeW(path.GetBuffer(maxPath), NULL, NULL, URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8);
	path.ReleaseBuffer();
	if (!SUCCEEDED(hr)) {
		ShowMessageBox(CString(L"Failed to get file path from URI: ") + uri.get(), L"Error");
		return false;
	}
	outPath = path;
	return true;
}

HRESULT CIssueWnd::NewWindowRequestHandler(ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args)
{
	args->put_Handled(TRUE);	// block new window (inform the request is processed, and no need to create default new window)

	CString path;
	unique_cotaskmem_string uri;
	args->get_Uri(&uri);
	WebView2GetLocalFilePathFromUri(uri, path);

	// check if the path is folder. if folder is dropped, add \issue.json to the path
	DWORD dwAttr = ::GetFileAttributes((LPCTSTR)path);
	if ((dwAttr != INVALID_FILE_ATTRIBUTES) && (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {	// フォルダがドロップされた場合
		path += _T("/_issue.json");
	}
	// check if the file is .json and show issue
	if (path.Right(5).CompareNoCase(L".json") != 0) {
		ShowMessageBox(CString(L"only .json file can be opened: ") + path, L"Invalid File");
		return S_OK;
	}
	// check if the file exists.
	dwAttr = ::GetFileAttributes((LPCTSTR)path);
	if (dwAttr == INVALID_FILE_ATTRIBUTES || (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {	// 指定ファイルがない場合
		ShowMessageBox(CString(L"File not found: ") + path, L"File not found");
		return S_OK;
	}

	m_pParent->AddTab(path);

	return S_OK;
}

HRESULT CIssueWnd::NavigationStartingHandler(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
{
	CString path;
	unique_cotaskmem_string uri;
	args->get_Uri(&uri);

	// 内部生成テキストだった場合の処理
	const CString internalText(L"data:text/html;");
	path = uri.get();
	if (path.Left(internalText.GetLength()) == internalText) {
		return S_OK;
	}

	// ローカルファイルだった場合の処理
	if (path.Left(LocalHost.GetLength()) == LocalHost) {
		args->put_Cancel(TRUE);

		// create path string
		UrlUnescape(path.GetBuffer(), NULL, NULL, URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8);
		path = path.Mid(LocalHost.GetLength());	// leave last '/'
		path.SetAt(0, L'\\');
		const int MAXPATH = 32767 + 1;
		CString localPath = m_IssueFilePath;

		if (path.Right(1) == L"/") {
			// a link from search result
			localPath = m_pParent->m_AppFolderPath;
			path = localPath + path + L"_issue.json";
		}
		else {
			// a link from issue
			localPath.Replace(L'/', L'\\');
			if (!SUCCEEDED(PathCchRemoveFileSpec(localPath.GetBuffer(MAXPATH), MAXPATH))) {
				localPath.ReleaseBuffer();
				ShowMessageBox(CString(L"PathCchRemoveFileSpec failed: ") + localPath, L"InternalError");
				return S_OK;
			}
			localPath.ReleaseBuffer();
			path = localPath + L"\\.." + path;
		}
		path.Replace(L'/', L'\\');

		// open file
		if (!PathFileExists(path)) {
			ShowMessageBox(CString(L"file does not exist: ") + path, L"Error");
			return S_OK;
		}
		if (path.Right(IssueFileName.GetLength()) == IssueFileName) {
			m_IssueFilePath = path;
			ShowIssue();
			return S_OK;
		}
		HINSTANCE hInst = ::ShellExecute(NULL, L"open", path, NULL, NULL, SW_SHOWNORMAL);
		if ((INT_PTR)hInst <= 32) {	// 戻り値が 32 より大きければ成功
			ShowMessageBox(CString(L"Failed to open file: ") + path, L"Error");
		}
		return S_OK;
	}

	WebView2GetLocalFilePathFromUri(uri, path);

	// issueファイルだった場合の処理
	if (path.Right(IssueFileName.GetLength() - 1) == IssueFileName.Mid(1)) {	// ignore '\\' on the head of IssueFileName
		m_IssueFilePath = path;
		ShowIssue();
		args->put_Cancel(TRUE);
		return S_OK;
	}

	// Issue.htmlだった場合の処理
	if (path.Right(IssueTemplate.GetLength() -1 ) == IssueTemplate.Mid(1)) {	// ignore '\\' on the head of IssueTemplate
		m_pParent->m_CtrlTabViewer.SetTabLabel(m_pParent->m_CtrlTabViewer.GetActiveTab(), L"Start");
		return S_OK;
	}

	args->put_Cancel(TRUE);

	return S_OK;
}

void CIssueWnd::ShowMessageBox(const CString& msg, const CString& title)
{
	if (m_pParent == NULL) {
		OutputDebugString(L"Error: m_pParent is NULL\n");
		MessageBox(msg, title, MB_OK | MB_ICONINFORMATION);
		return;
	}
	CString* wParam = new CString;
	CString* lParam = new CString;
	*wParam = msg;
	*lParam = title;
	m_pParent->PostMessageW(WM_USER_SHOW_MESSAGE_BOX, (WPARAM)wParam, (LPARAM)lParam);
}

bool CIssueWnd::ShowIssue()
{
	nlohmann::json json = m_pParent->ReadJson(m_IssueFilePath);
	if (json.empty()) {
		ShowMessageBox(L"Failed to open the file.", L"Error");
		return false;
	}
	if (!json.contains("issue")) {
		ShowMessageBox(L"This json file does not include issue", L"Error");
		return false;
	}

	// set localhost URL
	CString localhostURL;
	localhostURL.Format(L"%s/%d/", (LPCWSTR)LocalHost, json["issue"]["id"].get<int>());
	json["_localhost"] = std::string(CW2A(localhostURL)).c_str();

	// process header
	json["issue"]["description"] = ReplaceIssueLink(json["issue"]["description"].get<std::string>());
	json["issue"]["description"] = ConvertMdToHtml(json["issue"]["description"].get<std::string>());

	// process attachments
	if (json["issue"].contains("attachments")) {
		for (auto& attachment : json["issue"]["attachments"]) {
			CString filename = (LPCWSTR)CA2W(attachment["filename"].get<std::string>().c_str(), CP_UTF8);
			CString id;
			id.Format(L"%d", attachment["id"].get<int>());
			attachment["local_filename"] = std::string(CW2A((LPCWSTR)GetFileNameFromJson(filename, id), CP_UTF8));
		}
	}

	// process notes
	if (json["issue"].contains("journals")) {
		for (auto& journal : json["issue"]["journals"]) {
			// replace IDs to names
			if (journal.contains("details")) {
				for (auto& detail : journal["details"]) {
					// process attachments
					if (detail.contains("property") && detail["property"].get<std::string>() == "attachment") {
						CString filename;
						if (detail.contains("new_value") && detail["new_value"].is_string()) {
							filename = (LPCWSTR)CA2W(detail["new_value"].get<std::string>().c_str(), CP_UTF8);
						}
						else if (detail.contains("old_value") && detail["old_value"].is_string()) {
							filename = (LPCWSTR)CA2W(detail["old_value"].get<std::string>().c_str(), CP_UTF8);
						}
						CString fileID = (LPCWSTR)CA2W(detail["name"].get<std::string>().c_str(), CP_UTF8);
						detail["local_filename"] = std::string(CW2A((LPCWSTR)GetFileNameFromJson(filename, fileID), CP_UTF8));
					}
					if (!detail.contains("property") || !detail.contains("name")) {
						continue;
					}
					ReplaceId(detail, "attr", "status_id", m_pParent->m_Statuses);
					ReplaceId(detail, "attr", "assigned_to_id", m_pParent->m_Members);
					ReplaceId(detail, "attr", "tracker_id", m_pParent->m_Trackers);
					ReplaceId(detail, "attr", "priority_id", m_pParent->m_Priorities);
				}
			}

			// convert markdown in journal notes
			if (journal.contains("notes")) {
				journal["notes"] = ReplaceIssueLink(journal["notes"].get<std::string>());
				journal["notes"] = ConvertMdToHtml(journal["notes"].get<std::string>());
			}
		}
	}

	// JSON データを HTML テンプレートに埋め込む
	try {
		std::string injaOutput = m_pParent->m_InjaEnv.render(m_pParent->m_IssueTemplate, json);
		m_WebView->NavigateToString(CString(CA2W(injaOutput.c_str(), CP_UTF8)));
	}
	catch (const std::exception& e) {
		ShowMessageBox(CString(L"Failed to render HTML: ") + CString(e.what()), L"Error");
		return false;
	}

	SetTabTitle(m_IssueFilePath);
	return true;
}

CString CIssueWnd::SetTabTitle(CString file)
{
	// gather information to show in the tab title
	LPWSTR fileName;
	CString title;
	fileName = PathFindFileName(file);
	if (fileName == IssueFileName.Mid(1)) {	// _Issue.json
		// get folder name
		CString tmp = file;
		tmp.Replace(L'/', L'\\');
		PathCchRemoveFileSpec(tmp.GetBuffer(), tmp.GetLength() + 1);
		CString folder = PathFindFileName(tmp);
		int id = _ttoi(folder);
		if (id == 0) { // as ID:0 doesn't exist, assume this as conversion fail.
			title = L"Unknown";
		}
		else {
			title.Format(L"#%d", id);
		}
	}
	else if (fileName == IssueTemplate.Mid(1)) {	// Issue.html
		title = L"Start";
	}
	else {
		title = L"UNKNOWN";
	}

	m_pParent->m_CtrlTabViewer.SetTabLabel(m_pParent->m_CtrlTabViewer.GetActiveTab(), title);

	return title;
}

void CIssueWnd::ReplaceId(nlohmann::json& json, std::string property, std::string name, std::map<int, std::string> data)
{
	if (json["property"].get<std::string>() == property && json["name"].get<std::string>() == name) {
		if (json.contains("old_value") && data.contains(atoi(json["old_value"].get<std::string>().c_str()))) {
			json["old_value"] = data[atoi(json["old_value"].get<std::string>().c_str())];
		}
		if (json.contains("new_value") && data.contains(atoi(json["new_value"].get<std::string>().c_str()))) {
			json["new_value"] = data[atoi(json["new_value"].get<std::string>().c_str())];
		}
	}
}

std::string CIssueWnd::ReplaceIssueLink(std::string text)
{
	std::regex re(R"((^|\s)#([0-9]+)($|\s))");
	std::string fmt = std::string(R"($1<a href=")") + (LPCSTR)CW2A(LocalHost) + R"(/$2/_issue.json"> #$2 </a>$3)";
	std::string ret = std::regex_replace(text, re, fmt);
	return ret;
}

std::string CIssueWnd::ConvertMdToHtml(std::string mdText)
{
	std::string html;

	// GFM拡張（テーブル、タスクリスト、打ち消し線など）を有効化
	unsigned int parser_flags = MD_FLAG_STRIKETHROUGH | MD_FLAG_COLLAPSEWHITESPACE | MD_FLAG_TABLES;
	unsigned int render_flags = MD_HTML_FLAG_SKIP_UTF8_BOM;

	int result = md_html(mdText.c_str(), (MD_SIZE)mdText.length(),
		ConvertMdToHtmlSub, &html, parser_flags, render_flags);

	if (result != 0) {	// failed to convert markdown to HTML
		return mdText;
	}
	return html;
}

void CIssueWnd::ConvertMdToHtmlSub(const MD_CHAR* text, MD_SIZE size, void* userData)
{
	std::string* out = static_cast<std::string*>(userData);
	out->append(text, size);
}


void CIssueWnd::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	if (m_hWnd == NULL || m_WebViewController == NULL) {
		return;
	}

	CRect rect;
	m_pParent->m_CtrlTabViewer.GetWndArea(rect);
	m_WebViewController->put_Bounds(rect);
}

void CIssueWnd::OnShowWindow(BOOL bShow, UINT nStatus)
{
	CWnd::OnShowWindow(bShow, nStatus);

	if (m_WebViewController == NULL) {
		return;
	}

	SetWindowPos(&wndBottom, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	m_WebViewController->put_IsVisible(bShow);

	BOOL flag = FALSE;
	m_WebView->get_CanGoBack(&flag);
	m_pParent->m_CtrlButtonBack.EnableWindow(flag ? TRUE : FALSE);
	m_WebView->get_CanGoForward(&flag);
	m_pParent->m_CtrlButtonForward.EnableWindow(flag ? TRUE : FALSE);
}

void CIssueWnd::OnDestroy()
{
	if (m_WebViewController) {
		m_WebViewController->Close();
		m_WebViewController = nullptr;
	}
	if (m_WebView) {
		m_WebView = nullptr;
	}

	CWnd::OnDestroy();
}


void CIssueWnd::OnClose()
{
	auto& tab = m_pParent->m_CtrlTabViewer;
	int nIndex = tab.GetTabFromHwnd(m_hWnd);
	if (nIndex < 0) {
		ShowMessageBox(L"Failed to find tab to close", L"WARNING");
		return;
	}
	int nActiveTab = tab.GetActiveTab();
	if (nActiveTab == nIndex) {
		if (nIndex == 0) {
			if (tab.GetTabsNum() == 1) {	// there's only 1 tab
			}
			else {
				tab.SetActiveTab(1);
			}
		}
		else {
			tab.SetActiveTab(nIndex - 1);
		}
	}
	tab.RemoveTab(nIndex);

	// as the object is deleted by PostNcDestroy, cannot call OnClose() here.
}


void CIssueWnd::PostNcDestroy()
{
	CWnd::PostNcDestroy();

	delete this;
}

