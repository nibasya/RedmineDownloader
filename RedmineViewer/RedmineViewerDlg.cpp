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
#include "../Shared/CommonConfig.h"


using namespace Microsoft::WRL;
using namespace wil;

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
				// 環境作成に失敗した場合はログやエラーハンドリングを行ってください
				return envResult;
			}

			// 環境からコントローラを作成（ダイアログの HWND を渡す）
			env->CreateCoreWebView2Controller(
				this->GetSafeHwnd(),
				Microsoft::WRL::Callback<
					ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
					[this](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT
				{
					if (FAILED(controllerResult) || controller == nullptr)
					{
						// コントローラ作成失敗のハンドリング
						return controllerResult;
					}

					// COM ポインタを保持
					m_WebViewController = controller;
					// コントローラから ICoreWebView2 を取得
					ICoreWebView2* rawWebView = nullptr;
					if (SUCCEEDED(m_WebViewController->get_CoreWebView2(&rawWebView)) && rawWebView != nullptr)
					{
						m_WebView = rawWebView; // wil::com_ptr に引き渡し

						// ファイルのD&Dを禁止する
						wil::com_ptr<ICoreWebView2Controller4> controller4 = m_WebViewController.try_query<ICoreWebView2Controller4>();
						if (controller4) {
							controller4->put_AllowExternalDrop(FALSE);
						}

						// ダイアログのクライアント領域に合わせて Bounds を設定
						CRect rc;
						m_CtrlWebView.GetWindowRect(&rc);
						ScreenToClient(&rc);
						RECT bounds = { rc.left, rc.top, rc.right, rc.bottom };
						m_WebViewController->put_Bounds(bounds);

						// 初期ページをロード
						TCHAR szPath[MAX_PATH];
						std::basic_string<TCHAR> tgtPath(L"file:///");
						GetCurrentDirectory(MAX_PATH, szPath);
						tgtPath.append(szPath);
						std::replace(tgtPath.begin(), tgtPath.end(),L'\\', L'/');
						tgtPath.append(L"/Issue.html");
						m_WebView->Navigate(tgtPath.c_str());
					}

					return S_OK;
				}).Get());
			return S_OK;
		}).Get());

	// hr をチェックして必要ならログ出力（ここでは無視）
	UNREFERENCED_PARAMETER(hr);
	
	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
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
		for (int counter = 0; counter < 4; counter++) {	// フォルダがロックされている可能性があるため、最大10回まで再試行}
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
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
}

void CRedmineViewerDlg::OnBnClickedButtonForward()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
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
		m_Env.set_html_autoescape(true); // HTML エスケープを有効にする
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
	// replace IDs to names
	if (json["issue"].contains("journals")) {
		for (auto& journal : json["issue"]["journals"]) {
			if (journal.contains("details")) {
				for (auto& detail : journal["details"]) {
					if (!detail.contains("property") || !detail.contains("name")) {
						continue;
					}
					ReplaceId(detail, "attr", "status_id", m_Statuses);
					ReplaceId(detail, "attr", "assigned_to_id", m_Members);
					ReplaceId(detail, "attr", "tracker_id", m_Trackers);
					ReplaceId(detail, "attr", "priority_id", m_Priorities);
				}
			}
		}
	}

	// JSON データを HTML テンプレートに埋め込む
	try {
		std::string renderedHtml = m_Env.render(m_IssueTemplate, json);
		m_WebView->NavigateToString(CString(CA2W(renderedHtml.c_str(), CP_UTF8)));
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

std::string UtcToLocal(std::string in, LPSYSTEMTIME pLocal)
{
	SYSTEMTIME utc;
	ZeroMemory(&utc, sizeof(utc));
	if (sscanf_s(in.c_str(), "%4hd-%2hd-%2hdT%2hd:%2hd:%2hdZ", &utc.wYear, &utc.wMonth, &utc.wDay, &utc.wHour, &utc.wMinute, &utc.wSecond) != 6) {
		return std::string("not UTC!");
	}
	if (SystemTimeToTzSpecificLocalTime(NULL, &utc, pLocal) == 0) {
		std::ostringstream out;
		out << "UTC->Local failed!" << (LPCSTR)(CW2A(GetLastErrorToString(), CP_UTF8));
		return out.str();
	}
	return "";
}

void CRedmineViewerDlg::SetupCallbacks()
{
	m_Env.add_callback("UtcToLocal", 1, CallbackUtcToLocal);

	m_Env.add_callback("UtcToAgo", 1, CallbackUtcToAgo);

	m_Env.add_callback("UtcToYMD", 1, CallbackUtcToYMD);
}

inja::json CallbackUtcToLocal(inja::Arguments& args)
{
	std::string in = args.at(0)->get<std::string>();
	SYSTEMTIME local;
	std::string err;
	err = UtcToLocal(in, &local);
	if (!err.empty()) {
		return err;
	}
	CStringA out;
	out.Format("%04d/%02d/%02d %02d:%02d:%02d", local.wYear, local.wMonth, local.wDay, local.wHour, local.wMinute, local.wSecond);
	return std::string((LPCSTR)out);

	//std::ostringstream out;
	//out << local.wYear << "/" << local.wMonth << "/" << local.wDay << " " << local.wHour << ":" << local.wMinute << ":" << local.wSecond;
	//return out.str();
//	return inja::json();
}

inja::json CallbackUtcToAgo(inja::Arguments& args)
{
	std::string in = args.at(0)->get<std::string>();
	SYSTEMTIME local;
	std::string err;
	err = UtcToLocal(in, &local);
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
	std::string in = args.at(0)->get<std::string>();
	SYSTEMTIME local;
	std::string err;
	err = UtcToLocal(in, &local);
	if (!err.empty()) {
		return err;
	}
	CStringA out;
	out.Format("%04d/%02d/%02d", local.wYear, local.wMonth, local.wDay);
	return std::string((LPCSTR)out);
}

