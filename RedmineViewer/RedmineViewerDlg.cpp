// RedmineViewerDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "RedmineViewer.h"
#include "RedmineViewerDlg.h"
#include "afxdialogex.h"
#include "CAboutDlg.h"
#include "RPTT.h"
#include "GetLastErrorToString.h"
#include <string>
#include <pathcch.h>
#include <regex>
#include "../Shared/CommonConfig.h"
#include "../Shared/CommonFunc.h"


using namespace Microsoft::WRL;
using namespace wil;

#pragma comment (lib, "pathcch")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CRedmineViewerDlg ダイアログ

CRedmineViewerDlg::CRedmineViewerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_REDMINEVIEWER_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRedmineViewerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BUTTON_LOAD, m_CtrlButtonLoad);
	DDX_Control(pDX, IDC_BUTTON_BACK, m_CtrlButtonBack);
	DDX_Control(pDX, IDC_BUTTON_FORWARD, m_CtrlButtonForward);
	DDX_Control(pDX, IDC_WEBVEW2, m_CtrlWebView);
	DDX_Control(pDX, IDC_BUTTON_RELOAD, m_CtrlButtonReload);
}

BEGIN_MESSAGE_MAP(CRedmineViewerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_LOAD, &CRedmineViewerDlg::OnBnClickedButtonLoad)
	ON_BN_CLICKED(IDC_BUTTON_BACK, &CRedmineViewerDlg::OnBnClickedButtonBack)
	ON_BN_CLICKED(IDC_BUTTON_FORWARD, &CRedmineViewerDlg::OnBnClickedButtonForward)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_DESTROY()
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDC_BUTTON_RELOAD, &CRedmineViewerDlg::OnBnClickedButtonReload)
END_MESSAGE_MAP()


// CRedmineViewerDlg メッセージ ハンドラー

BOOL CRedmineViewerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// "バージョン情報..." メニューをシステム メニューに追加します。

	// IDM_ABOUTBOX は、システム コマンドの範囲内になければなりません。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// このダイアログのアイコンを設定します。アプリケーションのメイン ウィンドウがダイアログでない場合、
	//  Framework は、この設定を自動的に行います。
	SetIcon(m_hIcon, TRUE);			// 大きいアイコンの設定
	SetIcon(m_hIcon, FALSE);		// 小さいアイコンの設定

	// TODO: 初期化をここに追加します。
	SetWindowText(CAboutDlg::GetAppVersion());	// バージョン情報の設定
	LoadSetting();
	SetupCallbacks();
	LoadCommonData();
	
	// COM初期化
	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) {
		MessageBox(L"COM initialization failed", L"Error", MB_ICONERROR);
		return FALSE;
	}

	// WebView2用一時フォルダ名の生成
	wchar_t tempPath[MAX_PATH];
	GetTempPath(MAX_PATH, tempPath);
	m_WebViewTempFolder = CString(tempPath) + L"RedmineViewerWebView2";

	// --- WebView2 初期化（WIL + WRL::Callback を利用） ---
	// CreateCoreWebView2EnvironmentWithOptions の完了コールバックを設定
	hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr, // browserExecutableFolder
		m_WebViewTempFolder, // userDataFolder
		nullptr, // environmentOptions (ICoreWebView2EnvironmentOptions* を渡す場合はここを変更)
		Microsoft::WRL::Callback<
			ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[this](HRESULT envResult, ICoreWebView2Environment* env) -> HRESULT
		{
			if (FAILED(envResult) || env == nullptr)
			{
				MessageBox(_T("Failed to create WebView2 environment"), _T("Error"), MB_ICONERROR);
				return envResult;
			}

			// Create controller and set up WebView2
			env->CreateCoreWebView2Controller(this->GetSafeHwnd(),
				Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
					[this](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
						return this->WebView2CreateController(result, controller);
					}).Get());
			return S_OK;
		}).Get());

	// hr をチェックして必要ならログ出力（ここでは無視）
	UNREFERENCED_PARAMETER(hr);
	
	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}


HRESULT CRedmineViewerDlg::WebView2CreateController(HRESULT result, ICoreWebView2Controller* controller)
{
	if (FAILED(result) || controller == nullptr) return result;

	m_WebViewController = controller;
	ICoreWebView2* rawWebView = nullptr;
	if (SUCCEEDED(m_WebViewController->get_CoreWebView2(&rawWebView)) && rawWebView != nullptr) {
		m_WebView = rawWebView;

		// configure bounds of the WebView2 control to fit the dialog's client area
		CRect rc;
		m_CtrlWebView.GetWindowRect(&rc);
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
					this->m_CtrlButtonBack.EnableWindow(flag ? TRUE : FALSE);
					webView->get_CanGoForward(&flag);
					this->m_CtrlButtonForward.EnableWindow(flag ? TRUE : FALSE);
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

bool CRedmineViewerDlg::WebView2GetLocalFilePathFromUri(const wil::unique_cotaskmem_string& uri, CString& outPath)
{
	// check if the URI is local file path
	CString path(uri.get());
	if (path.Left(8).CompareNoCase(_T("file:///")) == 0) {
		path = path.Mid(8);	// remove "file:///"
	}
	else {
		MessageBox(CString(_T("Only opening local file is allowed: ")) + path, _T("Error"), MB_ICONERROR);
		return false;
	}
	const int maxPath = 32767 + 1;
	HRESULT hr = UrlUnescapeW(path.GetBuffer(maxPath), NULL, NULL, URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8);
	path.ReleaseBuffer();
	if (!SUCCEEDED(hr)) {
		MessageBox(CString(_T("Failed to get file path from URI: ")) + uri.get(), _T("Error"), MB_ICONERROR);
		return false;
	}
	outPath = path;
	return true;
}

HRESULT CRedmineViewerDlg::NewWindowRequestHandler(ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args)
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
		MessageBox(CString(_T("only .json file can be opened: ")) + path, _T("Invalid File"), MB_ICONERROR);
		return S_OK;
	}
	// check if the file exists.
	dwAttr = ::GetFileAttributes((LPCTSTR)path);
	if (dwAttr == INVALID_FILE_ATTRIBUTES || (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {	// 指定ファイルがない場合
		MessageBox(CString(_T("File not found: ")) + path, _T("File not found"), MB_ICONERROR);
		return S_OK;
	}

	m_IssueFilePath = path;
	ShowIssue();

	return S_OK;
}

HRESULT CRedmineViewerDlg::NavigationStartingHandler(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
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
			MessageBox(CString(L"PathCchRemoveFileSpec failed: ") + localPath);
			return S_OK;
		}
		localPath.ReleaseBuffer();
		path = localPath + L"\\.." + path;
		if (!PathFileExists(path)) {
			MessageBox(CString(L"file does not exist: "), path);
			return S_OK;
		}
		if (path.Right(IssueFileName.GetLength()) == IssueFileName) {
			m_IssueFilePath = path;
			ShowIssue();
			return S_OK;
		}
		HINSTANCE hInst = ::ShellExecute(NULL, L"open", path, NULL, NULL, SW_SHOWNORMAL);
		if ((INT_PTR)hInst <= 32) {	// 戻り値が 32 より大きければ成功
			MessageBox(CString(L"Failed to open file: ") + path);
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


void CRedmineViewerDlg::OnDestroy()
{
	m_WebViewController.reset();
	m_WebViewController = nullptr;
	m_WebView.reset();
	m_WebView = nullptr;

	CDialogEx::OnDestroy();

	// TODO: ここにメッセージ ハンドラー コードを追加します。
	SaveSetting();
	if (PathFileExists(m_WebViewTempFolder)) {	// WebView2のユーザーデータフォルダを削除
		SHFILEOPSTRUCT fileOp = { 0 };
		fileOp.wFunc = FO_DELETE;
		TCHAR from[MAX_PATH + 1];
		_tcscpy_s(from, m_WebViewTempFolder.GetString());
		from[m_WebViewTempFolder.GetLength() + 1] = L'\0'; // ダブルヌルで終わる必要がある
		fileOp.pFrom = from;
		fileOp.fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT;
		int foRet = 0;
		for (int counter = 0; counter < 4; counter++) {	// フォルダがロックされている可能性があるため、最大4回(2 sec.)まで再試行}
			foRet = SHFileOperation(&fileOp);	// this fuction cant't use GetLastError(), so check return value
			if (foRet == 0) {
				break;
			}
			Sleep(500);	// フォルダがロックされている可能性があるため、少し待ってから再試行
		}
		if(foRet != 0) {
			// エラー処理（必要に応じてログ出力など）
			CString errorMsg;
			errorMsg.Format(L"Failed to delete WebView2 temp folder %s. Error code: %d", fileOp.pFrom, foRet);
			MessageBox(errorMsg, L"Error", MB_ICONERROR);
		}
	}
	CoUninitialize();
}


void CRedmineViewerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// ダイアログに最小化ボタンを追加する場合、アイコンを描画するための
//  下のコードが必要です。ドキュメント/ビュー モデルを使う MFC アプリケーションの場合、
//  これは、Framework によって自動的に設定されます。

void CRedmineViewerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 描画のデバイス コンテキスト

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// クライアントの四角形領域内の中央
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// アイコンの描画
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// ユーザーが最小化したウィンドウをドラッグしているときに表示するカーソルを取得するために、
//  システムがこの関数を呼び出します。
HCURSOR CRedmineViewerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CRedmineViewerDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);

	if (m_hWnd == nullptr || m_WebViewController == nullptr)
		return;

	// ダイアログのクライアント領域に合わせて Bounds を設定
	CRect rc;
	m_CtrlWebView.GetWindowRect(&rc);
	ScreenToClient(&rc);
	RECT bounds = { rc.left, rc.top, rc.right, rc.bottom };
	m_WebViewController->put_Bounds(bounds);
}

void CRedmineViewerDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// TODO: ここにメッセージ ハンドラー コードを追加するか、既定の処理を呼び出します。

	CDialogEx::OnGetMinMaxInfo(lpMMI);
}


void CRedmineViewerDlg::OnBnClickedButtonLoad()
{
	// JSONファイルを選択するダイアログを表示
	CFileDialog dlg(TRUE, L"json", nullptr, OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, L"JSON Files (*.json)|*.json|All Files (*.*)|*.*||");
	if (dlg.DoModal() != IDOK) {
		return;
	}
	CString filePath = dlg.GetPathName();
	m_IssueFilePath = filePath;
	ShowIssue();
}

void CRedmineViewerDlg::OnBnClickedButtonBack()
{
	BOOL canGoBack = FALSE;
	if (!m_WebView) {
		return;
	}
	m_WebView->get_CanGoBack(&canGoBack);
	if (canGoBack) {
		m_WebView->GoBack();
	}
}

void CRedmineViewerDlg::OnBnClickedButtonForward()
{
	BOOL canGoForward = FALSE;
	if (!m_WebView) {
		return;
	}
	m_WebView->get_CanGoForward(&canGoForward);
	if (canGoForward) {
		m_WebView->GoForward();
	}
}

void CRedmineViewerDlg::OnBnClickedButtonReload()
{
	LoadCommonData();
	ShowIssue();
}

void CRedmineViewerDlg::OnDropFiles(HDROP hDropInfo)
{
	// ドロップされたファイルの数を取得
	UINT fileCount = DragQueryFile(hDropInfo, 0xFFFFFFFF, nullptr, 0);
	if (fileCount > 0) {
		// 最初のファイルのパスを取得
		wchar_t filePath[MAX_PATH];
		if (DragQueryFile(hDropInfo, 0, filePath, MAX_PATH) > 0) {
			// 拡張子が .json かどうかをチェック
			CString strFilePath(filePath);
			if (strFilePath.Right(5).CompareNoCase(L".json") != 0) {
				MessageBox(L"Please drop a JSON file.", L"Invalid File", MB_ICONERROR);
				return;
			}
			m_IssueFilePath = filePath;
			ShowIssue();
		}
	}

	CDialogEx::OnDropFiles(hDropInfo);
}

void CRedmineViewerDlg::LoadCommonData()
{
	try {
		m_IssueTemplate = m_Env.parse_template(L"Issue.html");
	}
	catch (const inja::InjaError& e) {	// injaの例外をcatchする
		std::ostringstream oss;
		oss << "Failed to load template: " << e.what() << "\nType: " << e.type << "\nMessage: " << e.message << "\nLocation: Line " << e.location.line << ", Column " << e.location.column;
		MessageBox(CString(oss.str().c_str()), L"Error", MB_ICONERROR);
		return;
	}

	// optional items
	nlohmann::json json;

	m_Members.clear();
	CString commonDataPath = dynamic_cast<CRedmineViewerApp*>(AfxGetApp())->m_AppFolderPath + RedmineDataFolder;
	json = ReadJson(commonDataPath + MembershipsFileName);
	if (json.contains("memberships")) {
		for (auto member : json["memberships"]) {
			if (member.contains("user")) {
				auto user = member["user"];
				if (user.contains("id") && user.contains("name")) {
					//					_RPTTN(_T("member ID:%d name: %s\n"), i["user"]["id"].get<int>(), (LPCTSTR)CA2W(i["user"]["name"].get<std::string>().c_str(), CP_UTF8));
					m_Members[user["id"].get<int>()] = user["name"].get<std::string>();
				}
			}
		}
	}

	m_Statuses.clear();
	json = ReadJson(commonDataPath + IssueStatusesFileName);
	if (json.contains("issue_statuses")) {
		for (auto status : json["issue_statuses"]) {
			if (status.contains("id") && status.contains("name")) {
				//				_RPTTN(_T("status ID:%d name: %s\n"), status["id"].get<int>(), (LPCTSTR)CA2W(status["name"].get<std::string>().c_str(), CP_UTF8));
				m_Statuses[status["id"].get<int>()] = status["name"].get<std::string>();
			}
		}
	}

	m_Trackers.clear();
	json = ReadJson(commonDataPath + TrackersFileName);
	if (json.contains("trackers")) {
		for (auto tracker : json["trackers"]) {
			if (tracker.contains("id") && tracker.contains("name")) {
				m_Trackers[tracker["id"].get<int>()] = tracker["name"].get<std::string>();
			}
		}
	}

	m_Priorities.clear();
	json = ReadJson(commonDataPath + IssuePrioritiesFileName);
	if (json.contains("issue_priorities")) {
		for (auto priority : json["issue_priorities"]) {
			if (priority.contains("id") && priority.contains("name")) {
				m_Priorities[priority["id"].get<int>()] = priority["name"].get<std::string>();
			}
		}
	}
}

nlohmann::json CRedmineViewerDlg::ReadJson(const wchar_t* filePath)
{
	try {
		std::ifstream ifs(filePath);
		if (!ifs.is_open()) {
			return nlohmann::json();
		}
		return nlohmann::json::parse(ifs);
	}
	catch (const std::exception& e) {
		MessageBox(CString(L"Failed to parse JSON: ") + filePath + L" Error: " + CString(e.what()), L"Error", MB_ICONERROR);
		return nlohmann::json();
	}
}

bool CRedmineViewerDlg::ShowIssue()
{
	nlohmann::json json = ReadJson(m_IssueFilePath);
	if (json.empty()) {
		MessageBox(L"Failed to open the file.", L"Error", MB_ICONERROR);
		return false;
	}
	if (!json.contains("issue")) {
		MessageBox(L"This json file does not include issue");
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
						else if (detail.contains("old_value") && detail["old_value"].is_string()){
							filename = (LPCWSTR)CA2W(detail["old_value"].get<std::string>().c_str(), CP_UTF8);
						}
						CString fileID = (LPCWSTR)CA2W(detail["name"].get<std::string>().c_str(), CP_UTF8);
						detail["local_filename"] = std::string(CW2A((LPCWSTR)GetFileNameFromJson(filename, fileID), CP_UTF8));
					}
					if (!detail.contains("property") || !detail.contains("name")) {
						continue;
					}
					ReplaceId(detail, "attr", "status_id", m_Statuses);
					ReplaceId(detail, "attr", "assigned_to_id", m_Members);
					ReplaceId(detail, "attr", "tracker_id", m_Trackers);
					ReplaceId(detail, "attr", "priority_id", m_Priorities);
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

void CRedmineViewerDlg::ReplaceId(nlohmann::json& json, std::string property, std::string name, std::map<int, std::string> data)
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

std::string CRedmineViewerDlg::ReplaceIssueLink(std::string text)
{
	std::regex re(R"((^|\s)#([0-9]+)($|\s))");
	std::string fmt = std::string(R"($1<a href=")") + (LPCSTR)CW2A(LocalHost) + R"(/$2/_issue.json"> #$2 </a>$3)";
	std::string ret = std::regex_replace(text, re, fmt);
	return ret;
}

std::string CRedmineViewerDlg::ConvertMdToHtml(std::string mdText)
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

void CRedmineViewerDlg::ConvertMdToHtmlSub(const MD_CHAR* text, MD_SIZE size, void* userData)
{
	std::string* out = static_cast<std::string*>(userData);
	out->append(text, size);
}

void CRedmineViewerDlg::LoadSetting()
{
	CWinApp* pApp = AfxGetApp();
	CString section;
	
	section = L"WindowPosition";
	CRect rc;
	rc.top = pApp->GetProfileInt(section, L"top", -1);
	rc.bottom = pApp->GetProfileInt(section, L"bottom", -1);
	rc.left = pApp->GetProfileInt(section, L"left", -1);
	rc.right = pApp->GetProfileInt(section, L"right", -1);
	if (rc.top != -1) {
		CRect rcWork;
		SystemParametersInfo(SPI_GETWORKAREA, 0, &rcWork, 0);
		if (rc.IntersectRect(rc, rcWork)) {
			MoveWindow(&rc);
		}
	}
}

void CRedmineViewerDlg::SaveSetting()
{
	CWinApp* pApp = AfxGetApp();
	CString section;

	section = L"WindowPosition";
	CRect rc;
	GetWindowRect(&rc);
	pApp->WriteProfileInt(section, L"top", rc.top);
	pApp->WriteProfileInt(section, L"bottom", rc.bottom);
	pApp->WriteProfileInt(section, L"left", rc.left);
	pApp->WriteProfileInt(section, L"right", rc.right);
}

std::string UtcToLocal(const inja::Arguments& args, LPSYSTEMTIME pLocal)
{
	std::string in = args.at(0)->get<std::string>();
	std::string err = args.at(1)->get<std::string>();

	// inがエラー文字列と同じ場合はエラーとみなす
	if (in == err) {
		return err;
	}

	SYSTEMTIME utc;
	ZeroMemory(&utc, sizeof(utc));
	if (sscanf_s(in.c_str(), "%4hd-%2hd-%2hdT%2hd:%2hd:%2hdZ", &utc.wYear, &utc.wMonth, &utc.wDay, &utc.wHour, &utc.wMinute, &utc.wSecond) != 6) {
		return "Not UTC! " + in;
	}
	if (SystemTimeToTzSpecificLocalTime(NULL, &utc, pLocal) == 0) {
		return std::string("UTC->Local failed!") + (LPCSTR)(CW2A(GetLastErrorToString(), CP_UTF8));
	}
	return "";
}

void CRedmineViewerDlg::SetupCallbacks()
{
	m_Env.add_callback("UtcToLocal", 2, CallbackUtcToLocal);

	m_Env.add_callback("UtcToAgo", 2, CallbackUtcToAgo);

	m_Env.add_callback("UtcToYMD", 2, CallbackUtcToYMD);
}

bool CallbackIs1stArgString(const inja::Arguments& args, inja::json& errText)
{
	const inja::json* in = args.at(0);
	const inja::json* err = args.at(1);

	if (!in->is_string()) {
		if (!err->is_string()) {
			if (in->is_null())
				errText = std::string("(null)");
			if (in->is_number())
				errText = std::string("(number)");
			if (in->empty())
				errText = std::string("(empty)");
		}
		else {
			errText = err->get<std::string>();
		}
		return false;
	}
	return true;
}

inja::json CallbackUtcToLocal(inja::Arguments& args)
{
	inja::json errText;
	if (!CallbackIs1stArgString(args, errText)) {
		return errText;
	}

	SYSTEMTIME local;
	std::string err = UtcToLocal(args, &local);
	if (!err.empty()) {
		return err;
	}

	CStringA out;
	out.Format("%04d/%02d/%02d %02d:%02d:%02d", local.wYear, local.wMonth, local.wDay, local.wHour, local.wMinute, local.wSecond);
	return std::string((LPCSTR)out);
}

inja::json CallbackUtcToAgo(inja::Arguments& args)
{
	inja::json errText;
	if (!CallbackIs1stArgString(args, errText)) {
		return errText;
	}

	SYSTEMTIME local;
	std::string err = UtcToLocal(args, &local);
	if (!err.empty()) {
		return err;
	}

	std::ostringstream oss;
	CTimeSpan ts = CTime::GetCurrentTime() - CTime(local);
	if (ts.GetDays() / 365 > 1) {
		oss << "over " << ts.GetDays() / 365 << " years";
	}
	else if (ts.GetDays() > 365) {
		oss << "1 year";
	}
	else if (ts.GetDays() / 30 > 1) {
		oss << ts.GetDays() / 30 << " months";
	}
	else if (ts.GetDays() > 30) {
		oss << "1 month";
	}
	else if (ts.GetDays() > 1) {
		oss << ts.GetDays() << " days";
	}
	else if (ts.GetDays() > 1) {
		oss << "1 day";
	}
	else if (ts.GetHours() > 1) {
		oss << ts.GetHours() << " hours";
	}
	else if (ts.GetHours() > 0) {
		oss << "1 hour";
	}
	else if (ts.GetMinutes() > 1) {
		oss << ts.GetMinutes() << " minutes";
	}
	else if (ts.GetMinutes() > 0) {
		oss << "1 minute";
	}
	else if (ts.GetSeconds() > 1) {
		oss << ts.GetSeconds() << " seconds";
	}
	else {
		oss << ts.GetSeconds() << " second";
	}
	return oss.str();
}

inja::json CallbackUtcToYMD(inja::Arguments& args)
{
	inja::json errText;
	if (!CallbackIs1stArgString(args, errText)) {
		return errText;
	}

	SYSTEMTIME local;
	std::string err = UtcToLocal(args, &local);
	if (!err.empty()) {
		return err;
	}

	CStringA out;
	out.Format("%04d/%02d/%02d", local.wYear, local.wMonth, local.wDay);
	return std::string((LPCSTR)out);
}

