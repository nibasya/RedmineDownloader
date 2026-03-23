#include "pch.h"
#include "framework.h"
#include "RedmineDownloader.h"
#include "RedmineDownloaderDlg.h"
#include "afxdialogex.h"
#include "CAboutDlg.h"
#include <string>
#include <sstream>
#include "rptt.h"
#include "GetLastErrorToString.h"
#include "../Shared/Common.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CRedmineDownloaderDlg ダイアログ

CRedmineDownloaderDlg::CRedmineDownloaderDlg(CWnd* pParent /*=nullptr*/)
    : CDialogEx(IDD_REDMINEDOWNLOADER_DIALOG, pParent), m_fStopThread(false), m_WorkerNew(0), m_WorkerUpdate(0), m_TargetInterval(0), m_fAutoExecute(false), m_TargetSaveVersionHistory(false)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRedmineDownloaderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_URL, m_CtrlEditUrl);
	DDX_Control(pDX, IDC_EDIT_API_KEY, m_CtrlEditApiKey);
	DDX_Control(pDX, IDC_EDIT_PROJECT_ID, m_CtrlEditProjectId);
	DDX_Control(pDX, IDC_EDIT_SAVETO, m_CtrlEditSaveTo);
	DDX_Control(pDX, IDC_BUTTON_SAVETO, m_CtrlButtonSaveTo);
	DDX_Control(pDX, IDC_EDIT_INTERVAL, m_CtrlEditInterval);
	DDX_Control(pDX, IDC_CHECK_SAVE_VERSION_HISTORY, m_CtrlCheckSaveVersionHistory);
	DDX_Control(pDX, IDC_EDIT_STATUS, m_CtrlEditStatus);
	DDX_Control(pDX, IDC_EDIT_TOTAL, m_CtrlEditTotal);
	DDX_Control(pDX, IDC_EDIT_NEW, m_CtrlEditNew);
	DDX_Control(pDX, IDC_EDIT_UPDATE, m_CtrlEditUpdate);
	DDX_Control(pDX, IDC_BUTTON_CANCEL, m_CtrlButtonCancel);
	DDX_Control(pDX, IDC_BUTTON_EXECUTE, m_CtrlButtonExecute);
	DDX_Control(pDX, IDC_CHECK_USE_CACHE, m_CtrlCheckUseCache);
	DDX_Control(pDX, IDC_BUTTON_CLEAR_CACHE, m_CtrlButtonClearCache);
}

BEGIN_MESSAGE_MAP(CRedmineDownloaderDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_EXECUTE, &CRedmineDownloaderDlg::OnBnClickedButtonExecute)
	ON_MESSAGE(WM_WORKER_STOPPED, &CRedmineDownloaderDlg::OnWorkerStopped)
	ON_MESSAGE(WM_WORKER_UPDATE_STATUS, &CRedmineDownloaderDlg::OnWorkerUpdateStatus)
	ON_MESSAGE(WM_WORKER_UPDATE_TOTAL, &CRedmineDownloaderDlg::OnWorkerUpdateTotal)
	ON_MESSAGE(WM_WORKER_UPDATE_NEW, &CRedmineDownloaderDlg::OnWorkerUpdateNew)
	ON_MESSAGE(WM_WORKER_UPDATE_UPDATE, &CRedmineDownloaderDlg::OnWorkerUpdateUpdate)
	ON_WM_DESTROY()
	ON_WM_GETMINMAXINFO()
	ON_BN_CLICKED(IDC_BUTTON_SAVETO, &CRedmineDownloaderDlg::OnBnClickedButtonSaveto)
	ON_BN_CLICKED(IDC_BUTTON_CANCEL, &CRedmineDownloaderDlg::OnBnClickedButtonCancel)
	ON_BN_CLICKED(IDC_BUTTON_CLEAR_CACHE, &CRedmineDownloaderDlg::OnBnClickedButtonClearCache)
END_MESSAGE_MAP()


// CRedmineDownloaderDlg メッセージ ハンドラー

BOOL CRedmineDownloaderDlg::OnInitDialog()
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
	GetWindowRect(&m_DefParentRect);
	SetWindowText(CAboutDlg::GetAppVersion());	// バージョン情報の設定
	LoadSettings();	// 設定の読み込み

	// 起動引数に -x が含まれている場合は自動実行フラグを立て、実行ボタンをクリックする
	CString cmdLine = AfxGetApp()->m_lpCmdLine;
	if (cmdLine.Find(_T("-x")) != -1) {
		m_fAutoExecute = true;
		// ダイアログ初期化後にコマンドをポストしてボタンクリックをシミュレート
		PostMessage(WM_COMMAND, MAKEWPARAM(IDC_BUTTON_EXECUTE, BN_CLICKED), (LPARAM)m_CtrlButtonExecute.GetSafeHwnd());
	}

	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}


void CRedmineDownloaderDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	SaveSettings();	// 設定の保存
}

void CRedmineDownloaderDlg::LoadSettings()
{
	CString section = _T("Settings");
	m_CtrlEditUrl.SetWindowText(theApp.GetProfileString(section, _T("Url"), _T("")));
	m_CtrlEditApiKey.SetWindowText(theApp.GetProfileString(section, _T("ApiKey"), _T("")));
	m_CtrlEditProjectId.SetWindowText(theApp.GetProfileString(section, _T("ProjectId"), _T("")));
	m_CtrlEditSaveTo.SetWindowText(theApp.GetProfileString(section, _T("SaveTo"), _T("")));
	m_CtrlEditInterval.SetWindowText(theApp.GetProfileString(section, _T("Interval"), _T("0")));
	m_CtrlCheckSaveVersionHistory.SetCheck(theApp.GetProfileInt(section, _T("SaveVersionHistory"), 0) ? BST_CHECKED : BST_UNCHECKED);
	m_CtrlCheckUseCache.SetCheck(theApp.GetProfileInt(section, _T("UseCache"), 0) ? BST_CHECKED : BST_UNCHECKED);
}

void CRedmineDownloaderDlg::SaveSettings()
{
	CString section = _T("Settings");
	CString buff;
	m_CtrlEditUrl.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("Url"), buff);
	m_CtrlEditApiKey.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("ApiKey"), buff);
	m_CtrlEditProjectId.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("ProjectId"), buff);
	m_CtrlEditSaveTo.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("SaveTo"), buff);
	m_CtrlEditInterval.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("Interval"), buff);
	theApp.WriteProfileInt(section, _T("SaveVersionHistory"), m_CtrlCheckSaveVersionHistory.GetCheck() == BST_CHECKED ? 1 : 0);
	theApp.WriteProfileInt(section, L"UseCache", m_CtrlCheckUseCache.GetCheck() == BST_CHECKED ? 1 : 0);
}

void CRedmineDownloaderDlg::EnableGui(bool fEnable)
{
	m_CtrlEditUrl.EnableWindow(fEnable);
	m_CtrlEditApiKey.EnableWindow(fEnable);
	m_CtrlEditProjectId.EnableWindow(fEnable);
	m_CtrlEditSaveTo.EnableWindow(fEnable);
	m_CtrlEditInterval.EnableWindow(fEnable);
	m_CtrlCheckSaveVersionHistory.EnableWindow(fEnable);
	m_CtrlCheckUseCache.EnableWindow(fEnable);
	m_CtrlButtonClearCache.EnableWindow(fEnable);
	m_CtrlButtonSaveTo.EnableWindow(fEnable);
	m_CtrlButtonExecute.EnableWindow(fEnable);
	m_CtrlButtonCancel.EnableWindow(!fEnable);
}


void CRedmineDownloaderDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	if (m_DefParentRect.Width() == 0)
		return;	// 初回呼び出し時は何もしない

	lpMMI->ptMinTrackSize = CPoint(m_DefParentRect.Width(), m_DefParentRect.Height());
	lpMMI->ptMaxTrackSize = CPoint(lpMMI->ptMaxTrackSize.x, m_DefParentRect.Height());

	CDialogEx::OnGetMinMaxInfo(lpMMI);
}


void CRedmineDownloaderDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CRedmineDownloaderDlg::OnPaint()
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
HCURSOR CRedmineDownloaderDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CRedmineDownloaderDlg::OnBnClickedButtonSaveto()
{
	IFileDialog* pDialog = NULL;
	HRESULT hr;
	DWORD options;
	CString path;

	// インスタンス生成
	hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDialog));
	if (FAILED(hr)) {
		MessageBox(TEXT("Failed to create dialog"));
		return;
	}

	// 設定の初期化
	pDialog->GetOptions(&options);
	pDialog->SetOptions(options | FOS_PICKFOLDERS);

	// フォルダ選択ダイアログを表示
	hr = pDialog->Show(NULL);

	// 結果取得
	if (!SUCCEEDED(hr)) {
		pDialog->Release();
		return;
	}
	IShellItem* pItem = NULL;
	PWSTR pPath = NULL;
	try {
		hr = pDialog->GetResult(&pItem);
		if (!SUCCEEDED(hr))
			throw hr;
		hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pPath);
		if (!SUCCEEDED(hr))
			throw hr;
		path = pPath;
		CoTaskMemFree(pPath);
		pItem->Release();
	}
	catch (HRESULT e) {
		if (e == S_OK) {
			OutputDebugString(_T("exception thrown with S_OK\n"));
		}
		if (pItem != NULL)
			pItem->Release();
		pDialog->Release();
		return;
	}
	pDialog->Release();

	m_CtrlEditSaveTo.SetWindowText(path);
	m_CtrlEditSaveTo.SetSel(0, -1, FALSE);
}


void CRedmineDownloaderDlg::OnBnClickedButtonExecute()
{
	CString buff;
	// ダイアログから保存のフォルダ, URL、APIキー、プロジェクトIDを取得
	m_CtrlEditSaveTo.GetWindowText(m_TargetFolder);	// 保存先入力欄から取得
	m_CtrlEditUrl.GetWindowText(m_TargetUrl); // URL入力欄から取得
    m_CtrlEditApiKey.GetWindowText(m_TargetApi); // APIキー入力欄から取得
    m_CtrlEditProjectId.GetWindowText(m_TargetProjectId); // プロジェクトID入力欄から取得
	m_CtrlEditInterval.GetWindowText(buff);	// インターバル入力欄から取得
	m_TargetInterval = _ttof(buff);

	if (m_TargetFolder.IsEmpty() || m_TargetUrl.IsEmpty() || m_TargetApi.IsEmpty() || m_TargetProjectId.IsEmpty())
	{
		MessageBox(_T("保存先、URL、APIキー、プロジェクトIDを入力してください。"));
		return;
	}

	m_TargetFolder = PrepareLongPath(m_TargetFolder);

	// 保存先フォルダの存在チェック
	DWORD dwAttr = ::GetFileAttributes((LPCTSTR)m_TargetFolder);
	if (dwAttr == INVALID_FILE_ATTRIBUTES || !(dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {
		MessageBox(_T("保存先フォルダが存在しません"));
		return;
	}

	// スレッド実行準備
	EnableGui(false); // GUIを無効化

	m_fStopThread = false; // ワーカースレッドの停止フラグをリセット
	m_cts = pplx::cancellation_token_source();	// cpprestのタスクキャンセルトークンを初期化
	m_CtrlEditStatus.SetWindowText(_T("スレッドを起動します"));
	if (AfxBeginThread(WorkerThread, (LPVOID)this) == NULL) {	// ワーカースレッドを開始
		OnWorkerStopped(0, 0);
	}

	return;
}


void CRedmineDownloaderDlg::OnBnClickedButtonCancel()
{
	m_cts.cancel();
	m_fStopThread = true; // ワーカースレッドの停止フラグをセット
}

LRESULT CRedmineDownloaderDlg::OnWorkerUpdateStatus(WPARAM wParam, LPARAM lParam)
{
	m_CtrlEditStatus.SetWindowText(m_WorkerStatus);
	return 0;
}

LRESULT CRedmineDownloaderDlg::OnWorkerUpdateTotal(WPARAM wParam, LPARAM lParam)
{
	m_CtrlEditTotal.SetWindowText(m_WorkerTotal);
	return 0;
}

LRESULT CRedmineDownloaderDlg::OnWorkerUpdateNew(WPARAM wParam, LPARAM lParam)
{
	CString buff;
	buff.Format(_T("%d"), m_WorkerNew);
	m_CtrlEditNew.SetWindowText(buff);
	return 0;
}

LRESULT CRedmineDownloaderDlg::OnWorkerUpdateUpdate(WPARAM wParam, LPARAM lParam)
{
	CString buff;
	buff.Format(_T("%d"), m_WorkerUpdate);
	m_CtrlEditUpdate.SetWindowText(buff);
	return 0;
}

UINT __cdecl CRedmineDownloaderDlg::WorkerThread(LPVOID pParam)
{
	CRedmineDownloaderDlg* pDlg = (CRedmineDownloaderDlg*)pParam;

	// スレッド開始をステータスに表示
	pDlg->m_WorkerStatus = _T("スレッド開始");
	::PostMessage(pDlg->m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	try {
		pDlg->CreateRedmineDataFolder();
		pDlg->GetMemberships();
		pDlg->GetStatuses();
		pDlg->GetTrackers();
		pDlg->GetPriorities();
		pDlg->GetIssueList();
		pDlg->UpdateLists();
		pDlg->GetIssue();
		pDlg->m_WorkerStatus = _T("完了");
		::PostMessage(pDlg->m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0); // 完了したことをステータスに表示
	}
	catch (CWorkerError e) {
		(void)e;	// remove C4101 unreferenced local variable warning
	}
	catch (const std::exception &e) {
		pDlg->m_WorkerStatus.Format(_T("エラー: %s"), (LPCTSTR)CString(e.what()));
		::PostMessage(pDlg->m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0); // 完了したことをステータスに表示
	}

	::PostMessage(pDlg->m_hWnd, WM_WORKER_STOPPED, 0, 0); // スレッド停止をメインスレッドに通知
	return 0;
}

void CRedmineDownloaderDlg::CreateRedmineDataFolder()
{
	DWORD dwAttr = ::GetFileAttributes((LPCTSTR)(m_TargetFolder + RedmineDataFolder));
	if ((dwAttr != INVALID_FILE_ATTRIBUTES) && (dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {
		return;	// すでに存在する場合は何もしない
	}
	if ((dwAttr != INVALID_FILE_ATTRIBUTES) && !(dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {
		MessageBox(L"保存先に RedmineData という名前のファイルが存在しています。保存先を変更するか、ファイルを削除してください。");
		throw CWorkerError();
	}
	// create directory
	if (CreateDirectory((LPCTSTR)(m_TargetFolder + RedmineDataFolder), NULL) == 0) {
		MessageBox(L"RedmineData フォルダの作成に失敗しました。保存先のアクセス権を確認してください。");
		throw CWorkerError();
	}
}

void CRedmineDownloaderDlg::DownloadFile(const CString& uri, web::http::http_response& response, const int issueNo)
{
	_RPTTN(_T("downloading %s\n"), uri);

	// HTTPクライアントを作成
	web::http::client::http_client client(utility::conversions::to_string_t((LPCTSTR)m_TargetUrl));
	// HTTPリクエストを作成
	web::http::http_request request(web::http::methods::GET);
	request.headers().add(L"X-Redmine-API-Key", utility::conversions::to_string_t((LPCTSTR)m_TargetApi));
	request.set_request_uri(utility::conversions::to_string_t((LPCTSTR)uri));
	// リクエストを送信
	response = client.request(request, m_cts.get_token()).get();
	// 成功したかチェック
	if (response.status_code() != web::http::status_codes::OK) {
		if (issueNo < 0) {
			m_WorkerStatus.Format(_T("Failed to get %s: %s"), (LPCTSTR)uri, (LPCTSTR)response.reason_phrase().c_str());
		}
		else {
			m_WorkerStatus.Format(_T("Issue %d: Failed to get %s: %s"), issueNo, (LPCTSTR)uri, (LPCTSTR)response.reason_phrase().c_str());
		}
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
	Sleep((DWORD)(m_TargetInterval * 1000));
}

void CRedmineDownloaderDlg::DownloadJson(const CString& uri, web::json::value& jsonResponse, const int issueNo)
{
	web::http::http_response response;
	DownloadFile(uri, response, issueNo);
	auto body = response.extract_string().get();

	// JSONのパース前に、禁止文字を除去しておく（Redmineのレスポンスに制御文字が混入することがあるため）
	utility::string_t cleaned;
	cleaned.reserve(body.size());

	for (auto ch : body)
	{
		// ここで ch は char または wchar_t（環境依存）
		// JSONで禁止される制御文字（0x00～0x1F）を削除（ただし \t \n \r は残す選択も可）
		// 今回は安全側に倒して \t \n \r 以外は除去する
		const auto code = static_cast<uint32_t>(ch);

		// 制御文字
		if (code <= 0x1F) {
			// JSONで許可される代表：タブ、LF、CR（必要なら残す）
			if (code == 0x09 || code == 0x0A || code == 0x0D) {
				cleaned.push_back(ch);
			}
			// それ以外は除去
			continue;
		}

		// DEL
		if (code == 0x7F) {
			continue;
		}

		// Unicode noncharacter を除去（U+FDD0..U+FDEF と、各面の U+FFFE/U+FFFF）
		// ただし ch が wchar_t の場合のみ確実に比較できる。charの場合はUTF-8の途中バイトなので
		// 「そのまま残す」方が安全だが、ここは簡易的に可能な範囲で除去する。
		if ((code >= 0xFDD0 && code <= 0xFDEF) ||
			(code & 0xFFFE) == 0xFFFE) {
			continue;
		}

		// 置換文字 U+FFFD を除去（デコード失敗の典型）
		if (code == 0xFFFD) {
			continue;
		}

		cleaned.push_back(ch);
	}

	jsonResponse = web::json::value::parse(cleaned);
}

void CRedmineDownloaderDlg::DownloadMultiPageJson(const CString& uri, web::json::value& jsonResponse, const utility::string_t key, const CString &option)
{
	const int initialValue = 10000000;
	int limit = initialValue;
	int totalCount = initialValue;
	CString defStatus = m_WorkerStatus;
	web::json::value integratedArray;
	web::json::value response;
	for (int offset = 0; offset < totalCount; offset += limit) {	// totalCount, limitは1ページ目のresponseから取得した値を使用
		// Update status message
		if (limit == initialValue) {
			// no status change;
		}
		else {
			m_WorkerStatus.Format(_T("%s %d/%d"), (LPCWSTR)defStatus, GetPage(offset + 1, limit), GetPage(totalCount, limit));
		}
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

		// limitやtotalCountの更新、ページ内容の取得
		CString completeUri;
		completeUri.Format(L"%s?&limit=%d&offset=%d%s", (LPCWSTR)uri, limit, offset, (LPCWSTR)option);
		DownloadJson(completeUri, response, -1);
		if (response.has_field(L"limit")) {
			limit = response[L"limit"].as_integer();	// responseからlimitを抽出
		}
		if (response.has_field(L"total_count")) {
			totalCount = response[L"total_count"].as_integer();	// responseからアイテム数を抽出
		}

		// extract target item
		auto targets = response[key];
		for (const auto& target : targets.as_array()) {
			integratedArray[integratedArray.size()] = target;	// 取得したtargetをIntegratedArray末尾に追加
		}
	}
	// integrate downloaded array into the latest downloaded json
	jsonResponse = response;
	jsonResponse[key] = integratedArray;
	if(jsonResponse.has_field(L"offset")) {
		jsonResponse[L"offset"] = web::json::value::number(0);	// 全データを含んでいるためoffsetを0にセット
	}
}

void CRedmineDownloaderDlg::SaveJson(const CString& saveTo, const web::json::value& jsonResponse, const CString& errorMessage)
{
	FILE* pFile = NULL;
	errno_t err = _wfopen_s(&pFile, saveTo, L"w");
	if (err != 0 || pFile == NULL) {
		if (!errorMessage.IsEmpty()) {
			m_WorkerStatus = errorMessage;
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		}
		else {
			m_WorkerStatus.Format(_T("JSONの保存に失敗: %s"), (LPCWSTR)saveTo);
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		}
		throw CWorkerError();
	}
	std::ostringstream oss;
	jsonResponse.serialize(oss);	// JSONをシリアル化 (UTF-8の文字列になる)
	//utility::string_t wstr = jsonResponse.serialize();	// JSONをシリアル化 (UTF-16の文字列になる)
	fwrite((WCHAR*)oss.str().c_str(), sizeof(char), oss.str().size(), pFile);	// JSONをファイルに保存
	fclose(pFile);
}

void CRedmineDownloaderDlg::GetMemberships()
{
	m_WorkerStatus = _T("メンバー一覧のダウンロード");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	web::json::value jsonResponse;
	CString uri;
	uri.Format(_T("/projects/%s/memberships.json"), (LPCTSTR)m_TargetProjectId);
	DownloadMultiPageJson(uri, jsonResponse, L"memberships");
	SaveJson(m_TargetFolder + RedmineDataFolder + MembershipsFileName, jsonResponse, _T("メンバー一覧の保存に失敗"));
}

void CRedmineDownloaderDlg::GetStatuses()
{
	m_WorkerStatus = _T("ステータス一覧のダウンロード");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	web::json::value jsonResponse;
	CString uri;
	uri.Format(_T("/issue_statuses.json"));
	DownloadMultiPageJson(uri, jsonResponse, L"issue_statuses");
	SaveJson(m_TargetFolder + RedmineDataFolder + IssueStatusesFileName, jsonResponse, _T("ステータス一覧の保存に失敗"));
}

void CRedmineDownloaderDlg::GetTrackers()
{
	m_WorkerStatus = _T("トラッカー一覧のダウンロード");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	web::json::value jsonResponse;
	CString uri;
	uri.Format(_T("/trackers.json"));
	DownloadMultiPageJson(uri, jsonResponse, L"trackers");
	SaveJson(m_TargetFolder + RedmineDataFolder + TrackersFileName, jsonResponse, _T("トラッカー一覧の保存に失敗"));
}

void CRedmineDownloaderDlg::GetPriorities()
{
	m_WorkerStatus = _T("優先度一覧のダウンロード");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	web::json::value jsonResponse;
	CString uri;
	uri.Format(_T("/enumerations/issue_priorities.json"));
	DownloadMultiPageJson(uri, jsonResponse, L"issue_priorities");
	SaveJson(m_TargetFolder + RedmineDataFolder + IssuePrioritiesFileName, jsonResponse, _T("優先度一覧の保存に失敗"));
}

void CRedmineDownloaderDlg::GetIssueList()
{
	try
	{
		web::json::value issueListJson;

		const int initialValue = 10000000;
		int limit = initialValue;
		int totalCount = initialValue;

		// Issue Listを取得する
		m_WorkerStatus.Format(_T("Issue一覧の取得"));	// postはDownloadMultiPageJson()が実施
		CString uri;
		uri.Format(L"/projects/%s/issues.json", (LPCWSTR)m_TargetProjectId);
		DownloadMultiPageJson(uri, issueListJson, L"issues", L"&status_id=*");
		totalCount = (int)issueListJson[L"issues"].size();	// responseからissue数を抽出
		m_WorkerTotal.Format(_T("%d"), totalCount);
		::PostMessageW(m_hWnd, WM_WORKER_UPDATE_TOTAL, 0, 0);

		// Issue Listを保存する
		m_WorkerStatus = _T("Issue一覧の保存");
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		CString saveTo(m_TargetFolder + IssueListFileName);	// 保存先のファイルパスを構築
		FILE* pFile = NULL;
		errno_t err = _wfopen_s(&pFile, saveTo, L"w");
		if (err != 0 || pFile == NULL) {
			m_WorkerStatus = _T("Issue一覧の保存に失敗");
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
			throw CWorkerError();
		}
		std::ostringstream oss;
		issueListJson.serialize(oss);	// JSONをシリアル化 (UTF-8の文字列になる)
		//utility::string_t wstr = issueListJson.serialize();	// JSONをシリアル化 (UTF-16の文字列になる)
		fwrite((WCHAR*)oss.str().c_str(), sizeof(char), oss.str().size(), pFile);	// JSONをファイルに保存
		fclose(pFile);

		m_WorkerStatus = _T("Issue一覧の保存完了");
	}
	catch (const web::http::http_exception e)
	{
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
}


void CRedmineDownloaderDlg::UpdateLists()
{
	m_WorkerStatus = _T("変化のあったIssueの抽出");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	// Issue一覧をファイルから読み込む
	CString path = m_TargetFolder + IssueListFileName;	// 保存先のファイルパスを構築
	web::json::value issueListJson;
	LoadJson(issueListJson, path);

	// cacheからデータを読み込む
	bool useCache = (m_CtrlCheckUseCache.GetCheck() == BST_CHECKED);
	web::json::value issueCacheJson;
	if (useCache) {
		CString cacheFile = m_TargetFolder + RedmineDataFolder + IssueListCacheFileName;
		if (PathFileExists(cacheFile)) {	// read cache only if it exists
			LoadJson(issueCacheJson, cacheFile);
		}
	}
	std::map<int, std::wstring> issueCache;
	if (issueCacheJson.is_null()) {
		useCache = false;
	}
	else {
		for (auto& p : issueCacheJson.as_array()) {
			issueCache[p[L"id"].as_integer()] = p[L"updated_on"].as_string();
		}
	}

	// Issue一覧から新しいIssueと更新が必要なIssueを抽出する
	m_NewIssue.clear();
	m_UpdateIssue.clear();
	auto issueObj = issueListJson[L"issues"];
	auto issueArray = issueObj.as_array();
	for (auto& issue : issueArray) {
		UINT issueID = issue[L"id"].as_integer();

		if (useCache) {
			// 新規issueか確認する
			auto it = issueCache.find(issueID);
			if (it == issueCache.end()) {
				m_NewIssue.push_back(issueID);
				continue;
			}

			// 更新が必要か確認する
			if (it->second != issue[L"updated_on"].as_string()) {
				m_UpdateIssue.push_back(issueID);
			}
		}
		else {
			CString issuePath;
			issuePath.Format(_T("%s\\%d"), (LPCTSTR)m_TargetFolder, issueID);	// Issueの保存先のファイルパスを構築
			DWORD dwAttr = ::GetFileAttributes((LPCTSTR)issuePath);
			if (dwAttr == INVALID_FILE_ATTRIBUTES || !(dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {
				m_NewIssue.push_back(issueID);
				continue;
			}

			// 更新が必要か確認する
			web::json::value currentIssue;
			CString issueJsonPath;
			issueJsonPath.Format(_T("%s\\%d%s"), (LPCTSTR)m_TargetFolder, issueID, (LPCTSTR)IssueFileName);	// IssueのJSONファイルのパスを構築
			// issue.jsonが存在しない場合は更新が必要とみなす
			if (PathFileExists((LPCTSTR)issueJsonPath) == FALSE) {
				m_UpdateIssue.push_back(issueID);
				continue;
			}
			// issue.jsonが存在する場合は、updated_onを比較して更新が必要か確認する
			LoadJson(currentIssue, issueJsonPath);	// IssueのJSONファイルを読み込む
			std::wstring updatedOnOld = currentIssue[L"issue"][L"updated_on"].as_string();
			std::wstring updatedOnNew = issue[L"updated_on"].as_string();
			if (updatedOnOld != updatedOnNew) {
				m_UpdateIssue.push_back(issueID);
			}
		}
	}
	m_WorkerNew = (int)m_NewIssue.size();
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_NEW, 0, 0);
	m_WorkerUpdate = (int)m_UpdateIssue.size();
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_UPDATE, 0, 0);
}

void CRedmineDownloaderDlg::LoadJson(web::json::value& jsonResponse, const CString& json)
{
	FILE* pFile = nullptr;
	errno_t err = _wfopen_s(&pFile, json, L"rb");
	if (err != 0 || pFile == nullptr) {
		m_WorkerStatus.Format(_T("JSONファイルの読み込みに失敗: %s"), (LPCTSTR)json);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}

	// ファイルサイズを取得
	if (_fseeki64(pFile, 0, SEEK_END) != 0) {
		fclose(pFile);
		m_WorkerStatus.Format(_T("JSONファイルの読み込みに失敗: %s"), (LPCTSTR)json);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
	long long fileSize = _ftelli64(pFile);
	if (fileSize < 0) fileSize = 0;
	_fseeki64(pFile, 0, SEEK_SET);

	std::vector<char> buffer((size_t)fileSize + 1);
	size_t readBytes = fread(buffer.data(), 1, (size_t)fileSize, pFile);
	buffer[readBytes] = '\0';
	fclose(pFile);

	try {
		jsonResponse = web::json::value::parse(utility::conversions::utf8_to_utf16(buffer.data()));
	}
	catch (const web::json::json_exception& e) {
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
}

void CRedmineDownloaderDlg::GetIssue()
{
	m_WorkerStatus = _T("Issueの保存");
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

	int i = 0;
	for (auto id : m_NewIssue) {
		i++;
		m_WorkerStatus.Format(_T("新規Issueの取得: %d/%d"), i, m_WorkerNew);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		GetIssue(id);
	}
	i = 0;
	for (auto id : m_UpdateIssue) {
		i++;
		m_WorkerStatus.Format(_T("更新されたIssueの取得: %d/%d"), i, m_WorkerUpdate);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		GetIssue(id);
	}
	SaveIssueListCache();
}

void CRedmineDownloaderDlg::GetIssue(UINT issueID)
{
	CString origMsg = m_WorkerStatus;
	CString issueSaveTo;
	try {
		m_WorkerStatus.Format(_T("%s downloading issue: %d"), (LPCTSTR)origMsg, issueID);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

		// IssueのJSONをダウンロード
		web::json::value jsonResponse;
		CString uri;
		uri.Format(L"/issues/%d.json?include=children,attachments,relations,changesets,journals,watchers", issueID);
		DownloadJson(uri, jsonResponse, issueID);

		// Issueの保存先のフォルダを作成
		CString issueFolder;
		issueFolder.Format(_T("%s\\%d"), (LPCTSTR)m_TargetFolder, issueID);	// Issueの保存先のファイルパスを構築
		if (!::CreateDirectory(issueFolder, NULL) && GetLastError() != ERROR_ALREADY_EXISTS) {
			m_WorkerStatus.Format(_T("Failed to create folder for issue: %d"), issueID);
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
			throw CWorkerError();
		}
		
		// IssueのJSONをファイルに保存
		issueSaveTo.Format(_T("%s\\%d%s"), (LPCTSTR)m_TargetFolder, issueID, (LPCTSTR)IssueFileName);	// 保存先のファイルパスを構築
		CString errorMessage;
		errorMessage.Format(_T("Failed to save issue: %d"), issueID);
		SaveJson(issueSaveTo, jsonResponse, errorMessage);

		// 添付ファイルのダウンロード
		auto& issue = jsonResponse[L"issue"];
		if(issue.has_field(L"attachments")) {	// 添付ファイルがあるか確認
			auto attachments = issue[L"attachments"].as_array();
			for (auto& attachment : attachments) {
				CString fileID;
				fileID.Format(L"%d", attachment[L"id"].as_integer()); // 添付ファイルのIDを取得
				CString fileUrl = attachment[L"content_url"].as_string().c_str();	// 添付ファイルのURLを取得
				CString fileUri = fileUrl.Mid(fileUrl.Find(L"/attachments/"));	// 添付ファイルのURIを取得
				CString fileName = attachment[L"filename"].as_string().c_str();	// 添付ファイルの名前を取得
				unsigned long long fileSize = attachment[L"filesize"].as_number().to_uint64(); // 添付ファイルのサイズを取得

				CString savePath;
				savePath.Format(_T("%s\\%d\\%s"), (LPCTSTR)m_TargetFolder, issueID, (LPCTSTR)GetFileNameFromJson(fileName, fileID));	// 添付ファイルの保存先のファイルパスを構築
				CFileFind finder;
				if (finder.FindFile(savePath)) {
					finder.FindNextFile();
					if (finder.GetLength() == fileSize) {	// すでに同名・同サイズのファイルが存在する場合はスキップ
						continue;
					}
				}

				m_WorkerStatus.Format(_T("%s downloading issue: %d file: %s"), (LPCTSTR)origMsg, issueID, (LPCTSTR)fileName);
				::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

				// 添付ファイルをダウンロードして保存する
				web::http::http_response fileResponse;
				DownloadFile(fileUri, fileResponse, issueID);
				auto bodyStream = fileResponse.body();
				auto outStream = concurrency::streams::fstream::open_ostream(utility::conversions::to_string_t((LPCTSTR)savePath)).get();
				bodyStream.read_to_end(outStream.streambuf()).get();
				outStream.close().get();
			}
		}

		if (m_TargetSaveVersionHistory) {
			// すでに存在する履歴を調査する
			CString historyPath;
			for (int count = 1; ; count++) {
				historyPath.Format(_T("%s\\%d\\%d.json"), (LPCTSTR)m_TargetFolder, issueID, count);
				if (!PathFileExists(historyPath)) {
					break;	// 履歴の保存先のファイルパスを構築して、存在しないファイルが見つかるまでループ
				}
			}
			// 新履歴として保存する
			CopyFile(issueSaveTo, historyPath, FALSE);
		}
	}
	catch (const web::http::http_exception e)
	{
		if (issueSaveTo.GetLength() > 0) {
			DeleteFile(issueSaveTo);	// エラーが発生した場合は保存したissue.jsonを削除する
		}
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
	catch (...) {
		if (issueSaveTo.GetLength() > 0) {
			DeleteFile(issueSaveTo);	// エラーが発生した場合は保存したissue.jsonを削除する
		}
		throw;
	}
}

void CRedmineDownloaderDlg::SaveIssueListCache()
{
	// Issue一覧をファイルから読み込む
	CString path = m_TargetFolder + IssueListFileName;	// 保存先のファイルパスを構築
	web::json::value issueListJson, issueCache;
	LoadJson(issueListJson, path);

	int count = 0;
	for (auto& issue : issueListJson[L"issues"].as_array()) {
		issueCache[count][L"id"] = issue[L"id"];
		issueCache[count][L"updated_on"] = issue[L"updated_on"];
		count++;
	}

	// IssueのIDとupdated_onのペアをキャッシュに保存する
	CString cacheFile = m_TargetFolder + RedmineDataFolder + IssueListCacheFileName;
	utility::string_t cacheString = issueCache.serialize();
	std::string utf8String = utility::conversions::to_utf8string(cacheString);
	std::ofstream ofs(cacheFile);
	if (ofs.is_open()) {
		ofs << utf8String;
		ofs.close();
	}

}

LRESULT CRedmineDownloaderDlg::OnWorkerStopped(WPARAM wParam, LPARAM lParam)
{
	EnableGui(true);	// GUIを有効化

    // 本関数はエラー発生時にも呼び出されるため、ここではStatus表示を更新しない
    // 自動実行モードならワーカ停止時にアプリケーションを終了する
    if (m_fAutoExecute) {
        m_fAutoExecute = false;
        PostMessage(WM_CLOSE);
    }
    // 必要なら他の UI 更新やログ処理をここに追加
    return 0;
}

CString CRedmineDownloaderDlg::PrepareLongPath(CString path) {
	// 1. すでにプレフィックスがある場合はそのまま返す
	if (path.Left(4) == _T("\\\\?\\")) {
		return (LPCTSTR)path;
	}

	// 2. 絶対パスを取得する
	// MAX_PATHを超えても取得できるよう、大きめのバッファを確保
	const int bufferSize = 32768;
	TCHAR fullPathBuffer[bufferSize];
	if (::GetFullPathName(path, bufferSize, fullPathBuffer, NULL) == 0) {
		return (LPCTSTR)path; // エラー時は元のパスを返す
	}

	CString fullPath(fullPathBuffer);

	// 3. 長いパス用のプレフィックスを付与
	CString longPath;
	if (fullPath.Left(2) == _T("\\\\")) {
		// UNCパス (\\server\share) の場合
		// -> \\?\UNC\server\share
		longPath.Format(_T("\\\\?\\UNC\\%s"), (LPCTSTR)fullPath.Mid(2));
	}
	else {
		// ローカルパス (C:\folder) の場合
		// -> \\?\C:\folder
		longPath.Format(_T("\\\\?\\%s"), (LPCTSTR)fullPath);
	}

	return longPath;
}

void CRedmineDownloaderDlg::OnBnClickedButtonClearCache()
{
	CString targetFolder;
	m_CtrlEditSaveTo.GetWindowText(targetFolder);
	CString cacheFile = targetFolder + RedmineDataFolder + IssueListCacheFileName;
	if (DeleteFile(cacheFile)) {
		m_CtrlEditStatus.SetWindowTextW(L"Cleared cache");
	}
	else {
		m_CtrlEditStatus.SetWindowTextW(CString(L"Failed to clear cache: ") + (LPCWSTR)GetLastErrorToString());
	}
}
