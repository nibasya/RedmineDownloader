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
#include "CIssueWnd.h"
#include <string>
#include <pathcch.h>
#include <future>
#include <atomic>
#include "../Shared/Common.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CRedmineViewerDlg ダイアログ

CRedmineViewerDlg::CRedmineViewerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_REDMINEVIEWER_DIALOG, pParent), m_TabViewWndId(2026)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRedmineViewerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BUTTON_LOAD, m_CtrlButtonLoad);
	DDX_Control(pDX, IDC_BUTTON_BACK, m_CtrlButtonBack);
	DDX_Control(pDX, IDC_BUTTON_FORWARD, m_CtrlButtonForward);
	DDX_Control(pDX, IDC_BUTTON_RELOAD, m_CtrlButtonReload);
	DDX_Control(pDX, IDC_BUTTON_RECACHE, m_CtrlButtonReCache);
	DDX_Control(pDX, IDC_EDIT_FIND, m_CtrlEditFind);
	DDX_Control(pDX, IDC_BUTTON_FIND, m_CtrlButtonFind);
}

BEGIN_MESSAGE_MAP(CRedmineViewerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_DESTROY()
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDC_BUTTON_LOAD, &CRedmineViewerDlg::OnBnClickedButtonLoad)
	ON_BN_CLICKED(IDC_BUTTON_BACK, &CRedmineViewerDlg::OnBnClickedButtonBack)
	ON_BN_CLICKED(IDC_BUTTON_FORWARD, &CRedmineViewerDlg::OnBnClickedButtonForward)
	ON_BN_CLICKED(IDC_BUTTON_RELOAD, &CRedmineViewerDlg::OnBnClickedButtonReload)
	ON_BN_CLICKED(IDC_BUTTON_RECACHE, &CRedmineViewerDlg::OnBnClickedButtonRecache)
	ON_BN_CLICKED(IDC_BUTTON_FIND, &CRedmineViewerDlg::OnBnClickedButtonFind)
	ON_MESSAGE(WM_USER_SHOW_MESSAGE_BOX, &CRedmineViewerDlg::OnShowMessageBox)
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
	// アプリの実行ファイルのパスを取得して、データ保存フォルダの基準にする
	TCHAR szPath[MAX_PATH];
	::GetModuleFileName(NULL, szPath, MAX_PATH);
	::PathRemoveFileSpec(szPath);
	m_AppFolderPath = szPath;

	HRESULT hr;

	GetWindowRect(&m_DefParentRect);
	SetWindowText(CAboutDlg::GetAppVersion());	// バージョン情報の設定
	LoadSetting();
	SetupCallbacks();
	LoadCommonData();
	LoadIssues(false);
	
	// COM初期化
	if(!AfxOleInit()){
		MessageBox(L"COM initialization failed", L"Error", MB_ICONSTOP);
		return FALSE;
	}

	// WebView2用一時フォルダ名の生成
	wchar_t tempPath[MAX_PATH];
	GetTempPath(MAX_PATH, tempPath);
	m_WebViewTempFolder = CString(tempPath) + L"RedmineViewerWebView2";

	// WebView2Environmentの生成
	hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr, // browserExecutableFolder
		m_WebViewTempFolder, // userDataFolder
		nullptr, // environmentOptions (ICoreWebView2EnvironmentOptions* を渡す場合はここを変更)
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[this](HRESULT envResult, ICoreWebView2Environment* env) -> HRESULT
			{
				if (FAILED(envResult) || env == nullptr)
				{
					MessageBox(L"Failed to create WebView2 environment", L"Error", MB_ICONSTOP);
					return envResult;
				}
				m_pWebEnvironment = env;
				return S_OK;
			}).Get());

	if (!SUCCEEDED(hr)) {
		MessageBox(L"Failed to create WebView2 environment", L"Error", MB_ICONSTOP);
		return false;
	}

	// CMFCTabCtrlの生成
	CRect rect;
	GetDlgItem(IDC_TAB_VIEWER)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	if (!m_CtrlTabViewer.Create(CMFCTabCtrl::STYLE_3D, rect, this, IDC_TAB_VIEWER, CMFCTabCtrl::LOCATION_TOP, WS_VISIBLE | WS_CHILD)) {
		MessageBox(L"Failed to create TabControl", L"Error", MB_ICONSTOP);
		return FALSE;
	}

	CRect clientRect;
	m_CtrlTabViewer.GetWndArea(clientRect);
	AddStartTab();
	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}


void CRedmineViewerDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	if (m_CtrlTabViewer.m_hWnd) {
		m_CtrlTabViewer.RemoveAllTabs();
	}

	SaveSetting();
	if (!m_WebViewTempFolder.IsEmpty() && PathFileExists(m_WebViewTempFolder)) {	// WebView2のユーザーデータフォルダを削除
		SHFILEOPSTRUCT fileOp = { 0 };
		fileOp.wFunc = FO_DELETE;
		TCHAR from[MAX_PATH + 1];
		_tcscpy_s(from, m_WebViewTempFolder.GetString());
		from[m_WebViewTempFolder.GetLength() + 1] = L'\0'; // ダブルヌルで終わる必要がある
		fileOp.pFrom = from;
		fileOp.fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT;
		int foRet = 0;
		for (int counter = 0; counter < 10; counter++) {	// フォルダがロックされている可能性があるため、最大10回(5 sec.)まで再試行}
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

	if (m_hWnd == NULL || m_CtrlTabViewer.m_hWnd == NULL)
		return;

	CRect rect;
	
	GetDlgItem(IDC_TAB_VIEWER)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	m_CtrlTabViewer.MoveWindow(&rect);
}

void CRedmineViewerDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	if (m_DefParentRect.Width() == 0)
		return;	// 初回呼び出し時は何もしない

	lpMMI->ptMinTrackSize = CPoint(m_DefParentRect.Width(), m_DefParentRect.Height());
	lpMMI->ptMaxTrackSize = CPoint(lpMMI->ptMaxTrackSize.x, lpMMI->ptMaxTrackSize.y);

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
	AddTab(filePath);
}

void CRedmineViewerDlg::OnBnClickedButtonBack()
{
	BOOL canGoBack = FALSE;
	auto webView = dynamic_cast<CIssueWnd*>(m_CtrlTabViewer.GetActiveWnd())->m_WebView;
	if (!webView) {
		return;
	}
	webView->get_CanGoBack(&canGoBack);
	if (canGoBack) {
		webView->GoBack();
	}
}

void CRedmineViewerDlg::OnBnClickedButtonForward()
{
	BOOL canGoForward = FALSE;
	auto webView = dynamic_cast<CIssueWnd*>(m_CtrlTabViewer.GetActiveWnd())->m_WebView;
	if (!webView) {
		return;
	}
	webView->get_CanGoForward(&canGoForward);
	if (canGoForward) {
		webView->GoForward();
	}
}

void CRedmineViewerDlg::OnBnClickedButtonReload()
{
	LoadCommonData();
	dynamic_cast<CIssueWnd*>(m_CtrlTabViewer.GetActiveWnd())->ShowIssue();
}

void CRedmineViewerDlg::OnDropFiles(HDROP hDropInfo)
{
	UINT fileCount = DragQueryFile(hDropInfo, 0xFFFFFFFF, nullptr, 0);
	if (fileCount > 0) {
		UINT maxPath = DragQueryFile(hDropInfo, 0, NULL, MAX_PATH) + 1;	// 最初のファイルのパスを取得
		CString filePath;
		if (DragQueryFile(hDropInfo, 0, filePath.GetBuffer(maxPath), maxPath) > 0) {
			filePath.ReleaseBuffer();
			// check if the path is folder. if folder is dropped, add \issue.json to the path
			DWORD dwAttr = ::GetFileAttributes((LPCTSTR)filePath);
			if ((dwAttr != INVALID_FILE_ATTRIBUTES) && (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {	// フォルダがドロップされた場合
				filePath += IssueFileName;
			}
			// 拡張子が .json かどうかをチェック
			if (filePath.Right(5).CompareNoCase(L".json") != 0) {
				MessageBox(L"Please drop a JSON file.", L"Invalid File", MB_ICONERROR);
				return;
			}
			AddTab(filePath);
		}
	}

	CDialogEx::OnDropFiles(hDropInfo);
}

void CRedmineViewerDlg::LoadCommonData()
{
	CString commonDataPath = m_AppFolderPath + RedmineDataFolder;
	try {
		m_IssueTemplate = m_InjaEnv.parse_template((LPCWSTR)(commonDataPath + IssueTemplate));
		m_SearchResultTemplate = m_InjaEnv.parse_template((LPCWSTR)(commonDataPath + SearchResultTemplate));
		m_ProjectTemplate = m_InjaEnv.parse_template((LPCWSTR)(commonDataPath + ProjectInfoTemplate));
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

nlohmann::json CRedmineViewerDlg::ReadJson(const wchar_t* filePath,bool suppressMessageBox)
{
	try {
		std::ifstream ifs(filePath);
		if (!ifs.is_open()) {
			if (!suppressMessageBox) {
				MessageBox(CString(L"Failed to open file: ") + filePath, L"Error", MB_ICONERROR);
			}
			else {
				throw std::exception("Failed to open file");	// メッセージボックスを表示しない場合は例外をスローする
			}
			return nlohmann::json();
		}
		return nlohmann::json::parse(ifs);
	}
	catch (const std::exception& e) {
		if (!suppressMessageBox) {
			MessageBox(CString(L"Failed to parse JSON: ") + filePath + L" Error: " + CString(e.what()), L"Error", MB_ICONERROR);
		}
		else {
			throw e;	// メッセージボックスを表示しない場合は例外を再スローする
		}
		return nlohmann::json();
	}
}

void CRedmineViewerDlg::AddStartTab()
{
	// check if issue list is loaded
	if (m_Issues.is_null() || !m_Issues.contains("issues") || m_Issues["issues"].empty()) {
		MessageBox(L"Failed to load issue list. Please check if the issue list cache is downloaded.", L"Error", MB_ICONERROR);
		return;
	}

	nlohmann::json project = m_Issues;
	// set project name
	project["project_name"] = m_Issues["issues"][0]["project"]["name"].get<std::string>();

	// set download time of issue list cache
	CString cacheName = m_AppFolderPath + RedmineDataFolder + IssueListCacheFileName;
	CFileStatus status;
	if (CFile::GetStatus(cacheName, status)) {
		CTime modifiedTime = status.m_mtime;
		project["download_on"] = modifiedTime.Format(L"%Y/%m/%d %H:%M:%S").GetString();
	}
	else {
		project["download_on"] = "Unknown";
	}

	// set last updated time from latest updated_on of issues
	auto it = std::max_element(m_Issues["issues"].begin(), m_Issues["issues"].end(), [](const nlohmann::json& a, const nlohmann::json& b) {
		std::string dateA = a["updated_on"].get<std::string>();
		std::string dateB = b["updated_on"].get<std::string>();
		return dateA < dateB;
		});
	if (it != m_Issues["issues"].end()) {
		project["updated_on"] = (*it)["updated_on"];
	} else {
		project["updated_on"] = "Unknown";
	}

	// set localhost URL
	CString localhostURL;
	localhostURL.Format(L"%s/", (LPCWSTR)LocalHost);
	project["_localhost"] = std::string(CW2A(localhostURL)).c_str();

	// JSON データを HTML テンプレートに埋め込む
	std::string injaOutput;
	try {
		injaOutput = m_InjaEnv.render(m_ProjectTemplate, project);
	}
	catch (const std::exception& e) {
		MessageBox(CString(L"Failed to render HTML: ") + CString(e.what()), L"Error", MB_ICONERROR);
		return;
	}
	if (injaOutput.size() == 0) {
		return;
	}

	AddTab((LPCWSTR)CA2W(injaOutput.c_str(), CP_UTF8));
}

void CRedmineViewerDlg::AddTab(CString file)
{
	// add tab
	if (m_TabViewWndId > 32760) {	// avoid overflow of the ID
		MessageBox(L"Cannot create tab anymore. Please restart the software.");
		return;
	}
	CIssueWnd* pNewWnd = new CIssueWnd(this);
	pNewWnd->CreateWebView2(m_TabViewWndId, file);
	m_TabViewWndId++;
	m_CtrlTabViewer.AddTab(pNewWnd, L"");
	m_CtrlTabViewer.SetActiveTab(m_CtrlTabViewer.GetTabsNum() - 1);
	m_CtrlTabViewer.EnableActiveTabCloseButton(TRUE);
}

void CRedmineViewerDlg::LoadSetting()
{
	CWinApp* pApp = AfxGetApp();
	CString section;

	section = L"WindowPosition";
	UINT nBytes = 0;
	LPBYTE pData = NULL;
	if (AfxGetApp()->GetProfileBinary(_T("Settings"), _T("WindowPlacement"), &pData, &nBytes))
	{
		if (nBytes == sizeof(WINDOWPLACEMENT))
		{
			SetWindowPlacement((WINDOWPLACEMENT*)pData);
		}
		delete[] pData;
	}
}

void CRedmineViewerDlg::SaveSetting()
{
	CWinApp* pApp = AfxGetApp();
	CString section;

	section = L"WindowPosition";
	WINDOWPLACEMENT wp;
	wp.length = sizeof(WINDOWPLACEMENT);
	if (GetWindowPlacement(&wp))
	{
		AfxGetApp()->WriteProfileBinary(_T("Settings"), _T("WindowPlacement"), (LPBYTE)&wp, sizeof(WINDOWPLACEMENT));
	}
}

void CRedmineViewerDlg::LoadIssues(bool reflesh)
{
	CString listFileName = m_AppFolderPath + RedmineDataFolder + IssueListFileName;
	CString cacheFileName = m_AppFolderPath + RedmineDataFolder + SearchCacheFileName;
	bool fListExists = PathFileExists(listFileName);
	bool fCacheExists = PathFileExists(cacheFileName);
	int iRet;

	m_Issues.clear();
	if (fCacheExists && !reflesh) {
		m_Issues = ReadJson(cacheFileName, false);
	}

	nlohmann::json updateList;
	nlohmann::json downloadedList;
	if(fListExists) {
		try {
			downloadedList = ReadJson(listFileName, false);
		}
		catch (const std::exception& e) {
			e;
			fListExists = false;
		}
	}
	if (!fListExists) {
		iRet = MessageBox(
			L"Failed to update cache as issue list generated by RedmineDownloader is not found or has error. Search issues from local folder?\n"
			L"OK: Search issues from local folder (might include deleted issues)\n"
			L"Cancel: skip updating issue cache",
			L"Failed", MB_OKCANCEL);
		if (iRet == IDCANCEL) {
			// IDCANCEL: skip updating issue cache
			if (reflesh) {
				MessageBox(L"As the search cache is refleshed, search is not available", L"Info", MB_ICONINFORMATION);
				return;
			}
			if (m_Issues.size() == 0) {
				MessageBox(L"As the search cache is empty, search is not available", L"Info", MB_ICONINFORMATION);
				return;
			}
			return;	// abort updating cache
		}
		// IDOK: search issues from local folder
		downloadedList = ScanLocalFolder();
	}
	updateList = CreateUpdateList(downloadedList);
	UpdateCache(updateList);
}

nlohmann::json CRedmineViewerDlg::ScanLocalFolder()
{
	int iRet;
	CFileFind finder;
	bool working = finder.FindFile(m_AppFolderPath + L"\\*.");
	bool skipError = false;	// if error occurs when reading file, do not show message box for each file and just skip the file.
	nlohmann::json updateList;
	while (working) {
		working = finder.FindNextFile();
		if (!finder.IsDirectory()) {
			continue;
		}
		if (finder.GetFilePath().Right(1) == L".") {	// skip "." and ".." folders
			continue;
		}
		CString issueFilePath = finder.GetFilePath() + IssueFileName;
		if (!PathFileExists(issueFilePath)) {
			continue;
		}
		nlohmann::json localIssue;
		try {
			localIssue = ReadJson(issueFilePath, true);
		}
		catch (const std::exception& e) {
			if (skipError) {
				continue;
			}
			iRet = MessageBox(CString(
				L"Failed to read file.\n"
				L"Yes: continue\n"
				L"No : continue, but never show error message\n"
				L"Cancel : stop reading\n"
				L"Filename : ") + issueFilePath
				+ L"\nError : " + CString(e.what()),
				L"Error", MB_YESNOCANCEL);
			if (iRet == IDNO) {
				skipError = true;
			}
			if (iRet == IDCANCEL) {
				MessageBox(L"Stop reading local files. Issue list may be incomplete.", L"Abort", MB_ICONINFORMATION);
				return updateList;
			}
			continue;
		}
		nlohmann::json item;
		item["id"] = localIssue["issue"]["id"];
		item["updated_on"] = localIssue["issue"]["updated_on"];
		updateList["issues"].push_back(item);
	}
	return updateList;
}

nlohmann::json CRedmineViewerDlg::CreateUpdateList(nlohmann::json& downloadedList)
{
	nlohmann::json updateList;
	// check ID difference of downloaded list and cached list
	auto compareById = [](const nlohmann::json& a, const nlohmann::json& b) {
		return a["id"].get<int>() < b["id"].get<int>();
		};
	std::sort(m_Issues["issues"].begin(), m_Issues["issues"].end(), compareById);
	std::sort(downloadedList["issues"].begin(), downloadedList["issues"].end(), compareById);
	std::set_difference(downloadedList["issues"].begin(), downloadedList["issues"].end(),
		m_Issues["issues"].begin(), m_Issues["issues"].end(),
		std::back_inserter(updateList["issues"]), compareById);
	// check updated items by comparing updated_on field
	std::vector<nlohmann::json> common;
	std::set_intersection(downloadedList["issues"].begin(), downloadedList["issues"].end(),
		m_Issues["issues"].begin(), m_Issues["issues"].end(),
		std::back_inserter(common), compareById);
	auto it = m_Issues["issues"].begin();
	for (const auto& item : common) {
		int id = item["id"].get<int>();
		auto itNew = std::find_if(it, m_Issues["issues"].end(), [id](const nlohmann::json& issue) {
			return issue["id"].get<int>() == id;
			});
		if (itNew != m_Issues["issues"].end()) {
			it = itNew;
			if (item["updated_on"] != (*it)["updated_on"]) {
				updateList["issues"].push_back(item);
			}
		}
	}
	return updateList;
}

void CRedmineViewerDlg::UpdateCache(const nlohmann::json& updateList)
{
	CDialog progDlg;
	progDlg.Create(IDD_ISSUE_LOAD_PROGRESS, this);
	progDlg.ShowWindow(SW_SHOW);
	auto& progCtrl = *(CProgressCtrl*)progDlg.GetDlgItem(IDC_PROGRESS);
	progCtrl.SetRange(0, (int)updateList["issues"].size());

	std::atomic<bool> isCancelled{ false };
	std::atomic<int> progress{ 0 };

	// worker loop
	auto future = std::async(std::launch::async, [&]() {
		bool skipError = false;	// if error occurs when reading file, do not show message box for each file and just skip the file.
		int iRet;
		for (auto& item : updateList["issues"]) {
			if (isCancelled) {
				return;
			}
			progress++;
			int id = item["id"].get<int>();
			CString file;
			file.Format(L"%d%s", id, (LPCWSTR)IssueFileName);
			if (!PathFileExists(file)) {
				if(skipError) {
					continue;
				}
				iRet = MessageBox(CString(L"Issue file: ") + file + L" is missing.\n"
					L"Preferred to execute RedmineDownloader.\n"
					L"Continue ? \n"
					L"Yes: continue\n"
					L"No: continue, and never show this message again\n"
					L"Continue: stop reading",
					L"Error", MB_YESNOCANCEL);
				if (iRet == IDCANCEL) {
					MessageBox(L"Cancelling cach update. Issue list may be incomplete.", L"Abort", MB_ICONINFORMATION);
					return;
				}
				if (iRet == IDNO) {
					skipError = true;
				}
				continue;
			}
			nlohmann::json issue;
			try {
				issue = ReadJson(file, true);
			} catch (const std::exception& e) {
				if(skipError) {
					continue;
				}
				iRet = MessageBox(CString(
					L"Failed to read file.\n"
					L"Yes: continue\n"
					L"No : continue, but never show error message\n"
					L"Cancel : stop reading\n"
					L"Filename : ") + file
					+ L"\nError : " + CString(e.what()),
					L"Error", MB_YESNOCANCEL);
				if (iRet == IDNO) {
					skipError = true;
				}
				if (iRet == IDCANCEL) {
					MessageBox(L"Cancelling cach update. Issue list may be incomplete.", L"Abort", MB_ICONINFORMATION);
					return;
				}
				continue;
			}
			m_Issues["issues"].push_back(issue["issue"]);
		}
		// save search cache
		CString cacheFile = m_AppFolderPath + RedmineDataFolder + SearchCacheFileName;
		std::ofstream ofs(cacheFile);
		if (ofs.is_open()) {
			ofs << m_Issues;
			ofs.close();
		}
	});

	// wait for load completion
	while (future.wait_for(std::chrono::milliseconds(100)) != std::future_status::ready) {
		_RPTTN(L"%d\n", (int)progress);
		progCtrl.SetPos(progress + 1 );
		progCtrl.SetPos(progress);
		if (!progCtrl.IsWindowVisible()) {
			isCancelled = true;
			break;
		}
		MSG msg;
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	progDlg.DestroyWindow();
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
	m_InjaEnv.add_callback("UtcToLocal", 2, CallbackUtcToLocal);
	m_InjaEnv.add_callback("UtcToAgo", 2, CallbackUtcToAgo);
	m_InjaEnv.add_callback("UtcToYMD", 2, CallbackUtcToYMD);
	m_InjaEnv.add_callback("LimitStr", 2, CallbackLimitStr);
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

inja::json CallbackLimitStr(inja::Arguments& args)
{
	CString in = (LPCWSTR)CA2W(args.at(0)->get<std::string>().c_str(), CP_UTF8);
	int length = args.at(1)->get<int>();
	CString ret = in.Left(length);
	if (ret.GetLength() == length) {
		ret += L"...";
	}
	return (LPCSTR)(CW2A(ret, CP_UTF8));
}


void CRedmineViewerDlg::OnBnClickedButtonRecache()
{
	LoadIssues(true);
}

void CRedmineViewerDlg::OnBnClickedButtonFind()
{
	// search
	std::vector<nlohmann::json> found;

	// prepare for search
	CString target;
	m_CtrlEditFind.GetWindowText(target);
	if (target.GetLength() == 0) {
		MessageBox(L"Please input search keyword", L"find", MB_ICONINFORMATION);
		return;
	}

	// check if the target is number for searching by ID
	LPWSTR endPtr;
	long searchId = _tcstol(target, &endPtr, 10);
	if (*endPtr == _T('\0') && !target.IsEmpty()) {	// check if the entire string was converted to a number
		nlohmann::json foundById;
		auto it = std::find_if(m_Issues["issues"].begin(), m_Issues["issues"].end(), [&](const nlohmann::json& b) {
			return b["id"].get<int>() == searchId;
			});
		if (it != m_Issues["issues"].end()) {	// found by ID
			foundById["issue"] = *it;
			foundById["found"][0]["string"] = (*it)["description"];
			found.push_back(foundById);
		}
	}
	// search by string
	SearchIssueList(found, m_Issues, (LPCSTR)CW2A((LPCWSTR)target, CP_UTF8));
	if (found.size() == 0) {
		MessageBox(L"Nothing found", L"find", MB_ICONINFORMATION);
		return;
	}

	// create json for output
	nlohmann::json result;
	const CString MARKER = L"__MARKER__";
	for (auto& p: found) {
		CString str = (LPCWSTR)ExtractContext((LPCWSTR)CA2W(p["found"][0]["string"].get<std::string>().c_str(), CP_UTF8), target, 100);
		str.Replace(target, MARKER);	// replace to MARKER for latter marking
		p["found"][0]["string"] = (LPCSTR)CW2A(str, CP_UTF8);
		result["result"].push_back(p);	// if possible, should have priority to show the search result; ex. title > description > latest note > other notes...
	}

	// set localhost URL
	CString localhostURL;
	localhostURL.Format(L"%s/", (LPCWSTR)LocalHost);
	result["_localhost"] = std::string(CW2A(localhostURL)).c_str();

	// set key
	result["key"] = (LPCWSTR)target;

	// JSON データを HTML テンプレートに埋め込む
	std::string injaOutput;
	try {
		injaOutput = m_InjaEnv.render(m_SearchResultTemplate, result);
		//		m_WebView->NavigateToString(CString(CA2W(injaOutput.c_str(), CP_UTF8)));
	}
	catch (const std::exception& e) {
		MessageBox(CString(L"Failed to render HTML: ") + CString(e.what()), L"Error", MB_ICONERROR);
		return;
	}

	if (injaOutput.size() == 0) {	// ShowSearchResult failed
		return;
	}

	// 検索ワードをハイライト
	CString html = (LPCWSTR)CA2W(injaOutput.c_str(), CP_UTF8);
	CString repl = CString(L"<mark>") + target + L"</mark>";
	html.Replace(MARKER, repl);

	AddTab(html);
}

LPARAM CRedmineViewerDlg::OnShowMessageBox(WPARAM wParam, LPARAM lParam)
{
	CString* msg = reinterpret_cast<CString*>(wParam);
	CString* title = reinterpret_cast<CString*>(lParam);
	
	MessageBox(*msg, *title, MB_ICONINFORMATION);

	delete msg;
	delete title;

	return 0;
}

void CRedmineViewerDlg::SearchIssueList(std::vector<nlohmann::json>& result, const nlohmann::json& issueList, const std::string& query)
{
	if (!issueList.contains("issues")) {
		MessageBox(L"Issue data cache is empty. Please update cache.", L"Error", MB_ICONERROR);
		return;
	}
	for (auto issue : issueList["issues"]) {
		nlohmann::json found;
		found["issue"] = issue;
		SearchIssue(found["found"], issue, nlohmann::json(), query);
		if (found["found"].size()) {
			found["issue"] = issue;
			result.push_back(found);
		}
	}
}

void CRedmineViewerDlg::SearchIssue(nlohmann::json& result, const nlohmann::json& issue, const nlohmann::json& givenKey, const std::string& query)
{
	if (issue.is_object()) {
		for (auto& [key, value] : issue.items()) {
			SearchIssue(result, value, key, query);
		}
	}
	else if (issue.is_array()) {
		for (size_t i = 0; i < issue.size(); ++i) {
			SearchIssue(result, issue[i], givenKey, query);
		}
	}
	else if (issue.is_string()) {
		if (issue.get<std::string>().find(query) != std::string::npos) {
			nlohmann::json tmp;
			tmp["string"] = issue;
			tmp["key"] = givenKey;
			result.push_back(tmp);
		}
	}
}

CString CRedmineViewerDlg::ExtractContext(const CString& str, const CString& key, int length)
{
	int index = str.Find(key);
	if (index == -1) {
		return _T("");
	}

	CString ret;
	int start = index - length/2;
	if (start < 0) {
		start = 0;
	}
	else {
		ret += L"...";
	}

	int remain = length - (index - start);
	int end = index + key.GetLength() + remain;
	if (end > str.GetLength()) {
		end = str.GetLength();
		return ret + str.Mid(start, end - start);
	}
	else {
		return ret + str.Mid(start, end - start) + L"...";
	}
}

void CRedmineViewerDlg::PostNcDestroy()
{
	m_pWebEnvironment = nullptr;

	CDialogEx::PostNcDestroy();
}
