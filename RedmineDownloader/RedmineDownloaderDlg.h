
// RedmineDownloaderDlg.h : ヘッダー ファイル
//

#pragma once

#include <list>
#include <cpprest/http_client.h>

// CRedmineDownloaderDlg ダイアログ
class CRedmineDownloaderDlg : public CDialogEx
{
// コンストラクション
public:
	CRedmineDownloaderDlg(CWnd* pParent = nullptr);	// 標準コンストラクター

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_REDMINEDOWNLOADER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV サポート


// 実装
protected:
	HICON m_hIcon;

	// 生成された、メッセージ割り当て関数
	virtual BOOL OnInitDialog();
	afx_msg void OnDestroy();
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonSaveto();
	afx_msg void OnBnClickedButtonExecute();
	afx_msg void OnBnClickedButtonCancel();
	afx_msg LRESULT OnWorkerStopped(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnWorkerUpdateStatus(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnWorkerUpdateTotal(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnWorkerUpdateNew(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnWorkerUpdateUpdate(WPARAM wParam, LPARAM lParam);
	CEdit m_CtrlEditUrl;
	CEdit m_CtrlEditApiKey;
	CEdit m_CtrlEditProjectId;
	CEdit m_CtrlEditSaveTo;
	CButton m_CtrlButtonSaveTo;
	CEdit m_CtrlEditInterval;
	CEdit m_CtrlEditStatus;
	CEdit m_CtrlEditTotal;
	CEdit m_CtrlEditNew;
	CEdit m_CtrlEditUpdate;

	CButton m_CtrlButtonExecute;
	CButton m_CtrlButtonCancel;
private:
	// RedmineDownloaderDlg自身の変数
	CRect m_DefParentRect;	// 親ウィンドウの初期サイズ


	// RedmineDownloaderDlg自身の関数
	void LoadSettings();		// 設定の読み込み
	void SaveSettings();		// 設定の保存
	void EnableGui(bool fEnable);	// GUIの有効/無効切り替え

	// ワーカースレッドの制御変数
	bool m_fStopThread;		// ワーカースレッドの停止用フラグ
	pplx::cancellation_token_source m_cts;	// cpprestの停止用トークン
	class CWorkerError {};	// ワーカースレッドを終了させるための例外クラス

	// ワーカースレッドとのI/F変数
	CString m_WorkerStatus;	// ワーカースレッドのステータス
	CString m_WorkerTotal;	// 総アイテム数
	CString m_TargetFolder;	// 保存先フォルダ名
	CString m_TargetUrl;	// RedmineのURL
	CString m_TargetApi;	// RedmineのAPIキー
	CString m_TargetProjectId;	// 対象プロジェクトのID
	double m_TargetInterval;	// 取得のインターバル（秒）
	int m_WorkerNew;	// 新アイテム数
	int m_WorkerUpdate;	// 更新アイテム数

	// ワーカースレッドの内部変数
	std::list<UINT> m_NewIssue;	// 新しいIssueの一覧
	std::list<UINT> m_UpdateIssue;	// アップデートが必要なIssueの一覧

	// ワーカースレッドの関数
	static UINT __cdecl WorkerThread(LPVOID pParam);	// ワーカースレッド
	void GetIssueList();	// Issue一覧の取得。エラー時にCWorkerErrorを投げる。
	void GetIssueListSub(int limit, int offset, web::http::http_response& response);
	void UpdateLists();		// json内のデータからnew issueやupdated issueを取り出す。エラー時にCWorkerErrorを投げる。
	void LoadJson(web::json::value& jsonResponse, const CString& json);	// jsonを読み込む。
	void GetIssue();		// 個別Issueを取得する。エラー時にCWorkerErrorを投げる。
	void GetIssue(UINT issueID);		// 個別Issueを取得する。エラー時にCWorkerErrorを投げる。
	CString PrepareLongPath(CString path);	// CString を受け取り、長いパス用に加工して返す関数
	int GetPage(int count, int limit) {
		return (count + limit - 1) / limit;
	};


};

// ワーカースレッドからCRedmineDownloaderDlgへの通知用メッセージ
const UINT WM_WORKER_STOPPED = WM_APP + 1;			// ワーカースレッドが停止することを通知
const UINT WM_WORKER_UPDATE_STATUS = WM_APP + 2;	// ワーカースレッドのステータス更新を通知
const UINT WM_WORKER_UPDATE_TOTAL = WM_APP + 3;	// ワーカースレッドの総アイテム数更新を通知
const UINT WM_WORKER_UPDATE_NEW = WM_APP + 4;	// ワーカースレッドの新アイテム数更新を通知
const UINT WM_WORKER_UPDATE_UPDATE = WM_APP + 5;	// ワーカースレッドの更新アイテム数更新を通知
