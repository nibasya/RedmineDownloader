// RedmineDownloaderDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "RedmineDownloader.h"
#include "RedmineDownloaderDlg.h"
#include "afxdialogex.h"

#include <cpprest/json.h>
#include <string>
#include <sstream>
#include "rptt.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

const CString IssueListFileName(_T("\\issue_list.json"));
const CString IssueFileName(_T("\\issue.json")); 

// アプリケーションのバージョン情報に使われる CAboutDlg ダイアログ

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV サポート

// 実装
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CRedmineDownloaderDlg ダイアログ



CRedmineDownloaderDlg::CRedmineDownloaderDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_REDMINEDOWNLOADER_DIALOG, pParent), m_fStopThread(false), m_TargetLimit(0), m_WorkerNew(0), m_WorkerUpdate(0)
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
	DDX_Control(pDX, IDC_EDIT_LIMIT, m_CtrlEditLimit);
	DDX_Control(pDX, IDC_EDIT_INTERVAL, m_CtrlEditInterval);
	DDX_Control(pDX, IDC_EDIT_STATUS, m_CtrlEditStatus);
	DDX_Control(pDX, IDC_EDIT_TOTAL, m_CtrlEditTotal);
	DDX_Control(pDX, IDC_EDIT_NEW, m_CtrlEditNew);
	DDX_Control(pDX, IDC_EDIT_UPDATE, m_CtrlEditUpdate);
	DDX_Control(pDX, IDC_BUTTON_CANCEL, m_CtrlButtonCancel);
	DDX_Control(pDX, IDC_BUTTON_EXECUTE, m_CtrlButtonExecute);
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

	LoadSettings();	// 設定の読み込み

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
	m_CtrlEditLimit.SetWindowText(theApp.GetProfileString(section, _T("Limit"), _T("0")));
	m_CtrlEditInterval.SetWindowText(theApp.GetProfileString(section, _T("Interval"), _T("0")));
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
	m_CtrlEditLimit.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("Limit"), buff);
	m_CtrlEditInterval.GetWindowText(buff);
	theApp.WriteProfileString(section, _T("Interval"), buff);
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
	m_CtrlEditLimit.GetWindowText(buff);	// 制限入力欄から取得
	m_TargetLimit = _ttoi(buff);

	if (m_TargetFolder.IsEmpty() || m_TargetUrl.IsEmpty() || m_TargetApi.IsEmpty() || m_TargetProjectId.IsEmpty() || m_TargetLimit == 0)
	{
		MessageBox(_T("保存先、URL、APIキー、プロジェクトID、制限値を入力してください。"));
		return;
	}

	// 保存先フォルダの存在チェック
	DWORD dwAttr = ::GetFileAttributes((LPCTSTR)m_TargetFolder);
	if (dwAttr == INVALID_FILE_ATTRIBUTES || !(dwAttr & FILE_ATTRIBUTE_DIRECTORY)) {
		MessageBox(_T("保存先フォルダが存在しません"));
		return;
	}

	// スレッド実行準備
	m_CtrlButtonExecute.EnableWindow(FALSE); // 実行ボタンを無効化して、二重実行を防止
	m_CtrlButtonCancel.EnableWindow(TRUE); // キャンセルボタンを有効化

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
		pDlg->GetIssueList();
		pDlg->UpdateLists();
		pDlg->GetIssue();
		pDlg->m_WorkerStatus = _T("完了");
		::PostMessage(pDlg->m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0); // 完了したことをステータスに表示
	}
	catch (CWorkerError e) {
		(void)e;	// remove C4101 unreferenced local variable warning
	}

	::PostMessage(pDlg->m_hWnd, WM_WORKER_STOPPED, 0, 0); // スレッド停止をメインスレッドに通知
	return 0;
}

void CRedmineDownloaderDlg::GetIssueList()
{
	try
	{
		web::http::http_response response;
		web::json::value jsonResponse;
		web::json::value issueListJson;

		// Issue数を取得する
		int totalCount = 0;

		m_WorkerStatus = _T("Issue数の取得");
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

		GetIssueListSub(1, 0, response);	// Issuesを取得
		jsonResponse = response.extract_json().get();
		totalCount = jsonResponse[L"total_count"].as_integer();	// responseからissue数を抽出
		m_WorkerTotal.Format(_T("%d"), totalCount);
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_TOTAL, 0, 0);

		// issueListの初期化
		issueListJson = jsonResponse;
		issueListJson.erase(L"issues");	// Issuesは後でまとめて取得するため、ここでは削除しておく

		web::json::value issueListArray = web::json::value::array();

		// Issue Listを取得する
		int endPage = (totalCount + m_TargetLimit - 1) / m_TargetLimit;
		for (int page = 1; page <= endPage; page++) {
			m_WorkerStatus.Format(_T("Issue一覧の取得 %d/%d"), page, endPage);
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

			GetIssueListSub(m_TargetLimit, m_TargetLimit * (page - 1), response);	// Issuesを取得
			jsonResponse = response.extract_json().get();
			auto issues = jsonResponse[L"issues"];
			for (const auto& issue : issues.as_array()) {
				issueListArray[issueListArray.size()] = issue;	// 取得したIssueをissueListArray末尾に追加
			}
		}

		m_WorkerStatus = _T("Issue一覧の保存");
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);

		issueListJson[L"issues"] = issueListArray;	// まとめたIssue ListをissueListJsonに追加
		CString saveTo(m_TargetFolder + IssueListFileName);	// 保存先のファイルパスを構築
		CStdioFile file;
		if (file.Open(saveTo, CFile::modeCreate | CFile::modeWrite)) {
			// UTF-8で保存したい場合
			std::ostringstream oss;
			issueListJson.serialize(oss);	// JSONをシリアル化 (UTF-8の文字列になる)
			file.Write((WCHAR*)oss.str().c_str(), (UINT)oss.str().size());	// JSONをファイルに保存

			// UTF-16で保存したい場合
			//utility::string_t wstr = issueListJson.serialize();	// JSONをシリアル化 (UTF-16の文字列になる)
			//file.Write((WCHAR*)wstr.c_str(), wstr.size()*sizeof(WCHAR));	// JSONをファイルに保存

			file.Close();
		} else {
			m_WorkerStatus = _T("Issue一覧の保存に失敗");
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
			throw CWorkerError();
		}
		m_WorkerStatus = _T("Issue一覧の保存完了");
	}
	catch (const web::http::http_exception e)
	{
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
}

void CRedmineDownloaderDlg::GetIssueListSub(int limit, int offset, web::http::http_response &response)
{
	// HTTPクライアントを作成
	web::http::client::http_client client(utility::conversions::to_string_t((LPCTSTR)m_TargetUrl));

	// リクエストURIを構築
	std::wostringstream requestURI;
	requestURI << L"/projects/" << (LPCTSTR)m_TargetProjectId << L"/issues.json?status_id=*&limit=" << limit << L"&offset=" << offset;

	// HTTPリクエストを作成
	web::http::http_request request(web::http::methods::GET);
	request.headers().add(L"X-Redmine-API-Key", utility::conversions::to_string_t((LPCTSTR)m_TargetApi));
	request.set_request_uri(requestURI.str());

	// リクエストを送信
	response = client.request(request, m_cts.get_token()).get();

	// 成功したかチェック
	if (response.status_code() != web::http::status_codes::OK) {
		m_WorkerStatus = _T("Failed to get issue lists");
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

	// Issue一覧から新しいIssueと更新が必要なIssueを抽出する
	m_NewIssue.clear();
	m_UpdateIssue.clear();
	auto issueObj = issueListJson[L"issues"];
	auto issueArray = issueObj.as_array();
	for (auto& issue : issueArray) {
		UINT issueID = issue[L"id"].as_integer();

		// 新規issueか確認する
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
		LoadJson(currentIssue, issueJsonPath);	// IssueのJSONファイルを読み込む
		auto val = currentIssue[L"issue"];
		std::wstring updatedOnOld = val[L"updated_on"].as_string();
		std::wstring updatedOnNew = issue[L"updated_on"].as_string();
		if (updatedOnOld != updatedOnNew) {
			m_UpdateIssue.push_back(issueID);
		}
	}
	m_WorkerNew = (int)m_NewIssue.size();
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_NEW, 0, 0);
	m_WorkerUpdate = (int)m_UpdateIssue.size();
	::PostMessage(m_hWnd, WM_WORKER_UPDATE_UPDATE, 0, 0);
}

void CRedmineDownloaderDlg::LoadJson(web::json::value& jsonResponse, const CString& json)
{
	CStdioFile file;
	if (!file.Open(json, CFile::modeRead)) {
		throw CWorkerError();
	}
	CFileStatus fileStat;
	file.GetStatus(fileStat);
	// UTF-8で保存した場合
	CHAR* buffer = new CHAR[(UINT)(fileStat.m_size / sizeof(CHAR)) + 1];	// ファイルサイズに基づいてバッファを確保
	file.Read(buffer, (UINT)fileStat.m_size);	// ファイルをバッファに読み込む
	buffer[fileStat.m_size / sizeof(CHAR)] = L'\0';	// バッファの末尾にNULL文字を追加
	try {
		jsonResponse = web::json::value::parse(utility::conversions::utf8_to_utf16((char*)buffer));	// バッファからJSONを解析
	}
	// UTF-16で保存した場合
	//WCHAR* buffer = new WCHAR[(UINT)(fileStat.m_size / sizeof(WCHAR)) + 1];	// ファイルサイズに基づいてバッファを確保
	//file.Read(buffer, (UINT)fileStat.m_size);	// ファイルをバッファに読み込む
	//buffer[fileStat.m_size / sizeof(WCHAR)] = L'\0';	// バッファの末尾にNULL文字を追加
	//try {
	//	jsonResponse = web::json::value::parse(buffer);	// バッファからJSONを解析
	//}
	catch (const web::json::json_exception e) {
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		delete buffer;
		throw CWorkerError();
	}
	delete buffer;
}

void CRedmineDownloaderDlg::GetIssue()
{
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
}

void CRedmineDownloaderDlg::GetIssue(UINT issueID)
{
	try {
		// Issueを取得する
		web::http::client::http_client client(utility::conversions::to_string_t((LPCTSTR)m_TargetUrl));
		// リクエストURIを構築
		std::wostringstream requestURI;
		requestURI << L"/issues/" << issueID << L".json?include=children,attachments,relations,changesets,journals,watchers";
		// HTTPリクエストを作成
		web::http::http_request request(web::http::methods::GET);
		request.headers().add(L"X-Redmine-API-Key", utility::conversions::to_string_t((LPCTSTR)m_TargetApi));
		request.set_request_uri(requestURI.str());
		// リクエストを送信
		web::http::http_response response = client.request(request, m_cts.get_token()).get();
		// 成功したかチェック
		if (response.status_code() != web::http::status_codes::OK) {
			m_WorkerStatus.Format(_T("Failed to get issue: %d"), issueID);
			throw CWorkerError();
		}

		// Issueの保存先のフォルダを作成
		CString issueFolder;
		issueFolder.Format(_T("%s\\%d"), (LPCTSTR)m_TargetFolder, issueID);	// Issueの保存先のファイルパスを構築
		if (!::CreateDirectory(issueFolder, NULL) && GetLastError() != ERROR_ALREADY_EXISTS) {
			m_WorkerStatus.Format(_T("Failed to create folder for issue: %d"), issueID);
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
			throw CWorkerError();
		}
		
		// IssueのJSONをファイルに保存
		web::json::value jsonResponse = response.extract_json().get();
		CString saveTo;
		saveTo.Format(_T("%s\\%d\\%s"), (LPCTSTR)m_TargetFolder, issueID, (LPCTSTR)IssueFileName);	// 保存先のファイルパスを構築
		CStdioFile file;
		if (file.Open(saveTo, CFile::modeCreate | CFile::modeWrite)) {
			// UTF-8で保存したい場合
			std::ostringstream oss;
			jsonResponse.serialize(oss);	// JSONをシリアル化 (UTF-8の文字列になる)
			file.Write((WCHAR*)oss.str().c_str(), (UINT)oss.str().size());	// JSONをファイルに保存
			// UTF-16で保存したい場合
			//utility::string_t wstr = jsonResponse.serialize();	// JSONをシリアル化 (UTF-16の文字列になる)
			//file.Write((WCHAR*)wstr.c_str(), wstr.size()*sizeof(WCHAR));	// JSONをファイルに保存
			file.Close();
		}
		else {
			m_WorkerStatus.Format(_T("Failed to save issue: %d"), issueID);
			::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
			throw CWorkerError();
		}
	}
	catch (const web::http::http_exception e)
	{
		m_WorkerStatus = CString(CA2W(e.what(), CP_UTF8));
		::PostMessage(m_hWnd, WM_WORKER_UPDATE_STATUS, 0, 0);
		throw CWorkerError();
	}
}

LRESULT CRedmineDownloaderDlg::OnWorkerStopped(WPARAM wParam, LPARAM lParam)
{
    // UI はメインスレッドで安全に更新する
    m_CtrlButtonExecute.EnableWindow(TRUE);
    m_CtrlButtonCancel.EnableWindow(FALSE);
//    m_CtrlEditStatus.SetWindowText(_T("停止しました"));	// 本関数はエラー発生時にも呼び出されるため、ここではStatus表示を更新しない
    // 必要なら他の UI 更新やログ処理をここに追加
    return 0;
}

