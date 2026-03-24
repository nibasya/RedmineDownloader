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
	HRESULT hr;
	GetWindowRect(&m_DefParentRect);
	SetWindowText(CAboutDlg::GetAppVersion());	// バージョン情報の設定
	LoadSetting();
	SetupCallbacks();
	LoadCommonData();
	
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
	AddTab();
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

void CRedmineViewerDlg::AddTab()
{
	// Load initial page
	TCHAR szPath[MAX_PATH];
	std::basic_string<TCHAR> tgtPath(L"file:///");
	GetCurrentDirectory(MAX_PATH, szPath);
	tgtPath.append(szPath);
	std::replace(tgtPath.begin(), tgtPath.end(), L'\\', L'/');
	tgtPath.append(IssueTemplate);
	AddTab(tgtPath.c_str());
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


void CRedmineViewerDlg::OnBnClickedButtonRecache()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
}

void CRedmineViewerDlg::OnBnClickedButtonFind()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
}

void CRedmineViewerDlg::PostNcDestroy()
{
	m_pWebEnvironment = nullptr;

	CDialogEx::PostNcDestroy();
}
