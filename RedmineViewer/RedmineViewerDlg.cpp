// RedmineViewerDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "RedmineViewer.h"
#include "RedmineViewerDlg.h"
#include "afxdialogex.h"
#include <wrl.h>
#include <wil/com.h>
#include <WebView2.h>


using namespace Microsoft::WRL;
using namespace wil;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


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
	// 	// COM初期化
	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) {
		MessageBox(L"COM initialization failed", L"Error", MB_ICONERROR);
		return FALSE;
	}

	// --- WebView2 初期化（WIL + WRL::Callback を利用） ---
	// CreateCoreWebView2EnvironmentWithOptions の完了コールバックを設定
	hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr, // browserExecutableFolder
		nullptr, // userDataFolder
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

						//// WebView の HWND を親ウィンドウにセット（既にコントローラが作成時に親設定されている場合もあります）
						//HWND hWndWebView = nullptr;
						//if (SUCCEEDED(m_WebViewController->get_Hwnd(&hWndWebView)) && hWndWebView != nullptr)
						//{
						//	::SetParent(hWndWebView, this->GetSafeHwnd());
						//}

						// ダイアログのクライアント領域に合わせて Bounds を設定
						CRect rc;
						m_CtrlWebView.GetWindowRect(&rc);
						ScreenToClient(&rc);
						RECT bounds = { rc.left, rc.top, rc.right, rc.bottom };
						m_WebViewController->put_Bounds(bounds);

						// 任意: 初期ページをロード
						m_WebView->Navigate(L"https://www.google.com");
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
	CDialogEx::OnDestroy();

	// TODO: ここにメッセージ ハンドラー コードを追加します。
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
	LoadJson(filePath);
}

void CRedmineViewerDlg::OnBnClickedButtonBack()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
}

void CRedmineViewerDlg::OnBnClickedButtonForward()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
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
			LoadJson(filePath);
		}
	}

	CDialogEx::OnDropFiles(hDropInfo);
}

bool CRedmineViewerDlg::LoadJson(const wchar_t* filePath)
{

}
