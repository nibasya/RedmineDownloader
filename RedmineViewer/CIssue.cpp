#include "pch.h"
#include "CIssue.h"
#include "RedmineViewerDlg.h"
#include "../Shared/CommonConfig.h"
#include <pathcch.h>
#include <regex>

using namespace Microsoft::WRL;
using namespace wil;

#pragma comment (lib, "pathcch")

bool CIssue::CreateWebView2()
{
	HRESULT hr;

	// --- WebView2 初期化（WIL + WRL::Callback を利用） ---
	// CreateCoreWebView2EnvironmentWithOptions の完了コールバックを設定
	hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr, // browserExecutableFolder
		m_pParent->m_WebViewTempFolder, // userDataFolder
		nullptr, // environmentOptions (ICoreWebView2EnvironmentOptions* を渡す場合はここを変更)
		Microsoft::WRL::Callback<
		ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[this](HRESULT envResult, ICoreWebView2Environment* env) -> HRESULT
			{
				if (FAILED(envResult) || env == nullptr)
				{
					m_pParent->MessageBox(L"Failed to create WebView2 environment", L"Error", MB_ICONERROR);
					return envResult;
				}

				// Create controller and set up WebView2
				env->CreateCoreWebView2Controller(m_pParent->m_CtrlTabViewer.m_hWnd,
					Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[this](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
							return this->WebView2CreateController(result, controller);
						}).Get());
				return S_OK;
			}).Get());

	// hr をチェックして必要ならログ出力（ここでは無視）
	UNREFERENCED_PARAMETER(hr);
}

HRESULT CIssue::WebView2CreateController(HRESULT result, ICoreWebView2Controller* controller)
{
	if (FAILED(result) || controller == nullptr) return result;

	m_WebViewController = controller;
	ICoreWebView2* rawWebView = nullptr;
	if (SUCCEEDED(m_WebViewController->get_CoreWebView2(&rawWebView)) && rawWebView != nullptr) {
		m_WebView = rawWebView;

		// configure bounds of the WebView2 control to fit the dialog's client area
		CRect rc;
		m_pParent->m_CtrlTabViewer.GetWindowRect(&rc);
		ScreenToClient(&rc);
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

		// Load initial page
		TCHAR szPath[MAX_PATH];
		std::basic_string<TCHAR> tgtPath(L"file:///");
		GetCurrentDirectory(MAX_PATH, szPath);
		tgtPath.append(szPath);
		std::replace(tgtPath.begin(), tgtPath.end(), L'\\', L'/');
		tgtPath.append(IssueTemplate);
		m_WebView->Navigate(tgtPath.c_str());
	}
	return S_OK;
}

bool CIssue::WebView2GetLocalFilePathFromUri(const wil::unique_cotaskmem_string& uri, CString& outPath)
{
	// check if the URI is local file path
	CString path(uri.get());
	if (path.Left(8).CompareNoCase(_T("file:///")) == 0) {
		path = path.Mid(8);	// remove "file:///"
	}
	else {
		m_pParent->MessageBox(CString(L"Only opening local file is allowed: ") + path, L"Error", MB_ICONERROR);
		return false;
	}
	const int maxPath = 32767 + 1;
	HRESULT hr = UrlUnescapeW(path.GetBuffer(maxPath), NULL, NULL, URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8);
	path.ReleaseBuffer();
	if (!SUCCEEDED(hr)) {
		m_pParent->MessageBox(CString(L"Failed to get file path from URI: ") + uri.get(), L"Error", MB_ICONERROR);
		return false;
	}
	outPath = path;
	return true;
}

HRESULT CIssue::NewWindowRequestHandler(ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args)
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
		m_pParent->MessageBox(CString(L"only .json file can be opened: ") + path, L"Invalid File", MB_ICONERROR);
		return S_OK;
	}
	// check if the file exists.
	dwAttr = ::GetFileAttributes((LPCTSTR)path);
	if (dwAttr == INVALID_FILE_ATTRIBUTES || (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {	// 指定ファイルがない場合
		m_pParent->MessageBox(CString(L"File not found: ") + path, L"File not found", MB_ICONERROR);
		return S_OK;
	}

	m_IssueFilePath = path;
	ShowIssue();

	return S_OK;
}

HRESULT CIssue::NavigationStartingHandler(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
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
		UrlUnescape(path.GetBuffer(), NULL, NULL, URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8);
		path = path.Mid(LocalHost.GetLength() - 1);	// leave last '/'
		path.SetAt(0, L'\\');
		const int MAXPATH = 32767 + 1;
		CString localPath = m_IssueFilePath;
		localPath.Replace(L'/', L'\\');
		path.Replace(L'/', L'\\');
		if (!SUCCEEDED(PathCchRemoveFileSpec(localPath.GetBuffer(MAXPATH), MAXPATH))) {
			localPath.ReleaseBuffer();
			m_pParent->MessageBox(CString(L"PathCchRemoveFileSpec failed: ") + localPath);
			return S_OK;
		}
		localPath.ReleaseBuffer();
		path = localPath + L"\\.." + path;
		if (!PathFileExists(path)) {
			m_pParent->MessageBox(CString(L"file does not exist: "), path);
			return S_OK;
		}
		if (path.Right(IssueFileName.GetLength()) == IssueFileName) {
			m_IssueFilePath = path;
			ShowIssue();
			return S_OK;
		}
		HINSTANCE hInst = ::ShellExecute(NULL, L"open", path, NULL, NULL, SW_SHOWNORMAL);
		if ((INT_PTR)hInst <= 32) {	// 戻り値が 32 より大きければ成功
			m_pParent->MessageBox(CString(L"Failed to open file: ") + path);
		}
		return S_OK;
	}
	WebView2GetLocalFilePathFromUri(uri, path);

	// issueファイルだった場合の処理
	if (path.Right(IssueFileName.GetLength()) == IssueFileName) {
		return E_NOTIMPL;
	}

	// Issue.htmlだった場合の処理
	if (path.Right(IssueTemplate.GetLength()) == IssueTemplate) {
		return S_OK;
	}

	args->put_Cancel(TRUE);

	return S_OK;
}

bool CIssue::ShowIssue()
{
	nlohmann::json json = m_pParent->ReadJson(m_IssueFilePath);
	if (json.empty()) {
		m_pParent->MessageBox(L"Failed to open the file.", L"Error", MB_ICONERROR);
		return false;
	}
	if (!json.contains("issue")) {
		m_pParent->MessageBox(L"This json file does not include issue");
		return false;
	}

	// set localhost URL
	CString localhostURL;
	localhostURL.Format(L"%s/%d/", LocalHost, json["issue"]["id"].get<int>());
	json["_localhost"] = std::string(CW2A(localhostURL));

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
		std::string injaOutput = m_Env.render(m_IssueTemplate, json);
		m_WebView->NavigateToString(CString(CA2W(injaOutput.c_str(), CP_UTF8)));
	}
	catch (const std::exception& e) {
		MessageBox(CString(L"Failed to render HTML: ") + CString(e.what()), L"Error", MB_ICONERROR);
		return false;
	}

	return true;
}

void CIssue::ReplaceId(nlohmann::json& json, std::string property, std::string name, std::map<int, std::string> data)
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

std::string CIssue::ReplaceIssueLink(std::string text)
{
	std::regex re(R"((^|\s)#([0-9]+)($|\s))");
	std::string fmt = std::string(R"($1<a href=")") + (LPCSTR)CW2A(LocalHost) + R"(/$2/_issue.json"> #$2 </a>$3)";
	std::string ret = std::regex_replace(text, re, fmt);
	return ret;
}

std::string CIssue::ConvertMdToHtml(std::string mdText)
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

void CIssue::ConvertMdToHtmlSub(const MD_CHAR* text, MD_SIZE size, void* userData)
{
	std::string* out = static_cast<std::string*>(userData);
	out->append(text, size);
}

