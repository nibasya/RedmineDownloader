
// RedmineViewerDlg.h : ヘッダー ファイル
//

#pragma once

// CRedmineViewerDlg ダイアログ
class CRedmineViewerDlg : public CDialogEx
{
// コンストラクション
public:
	CRedmineViewerDlg(CWnd* pParent = nullptr);	// 標準コンストラクター

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_REDMINEVIEWER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV サポート


// 実装
protected:
	HICON m_hIcon;

	// 生成された、メッセージ割り当て関数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CButton m_CtrlButtonLoad;
	CButton m_CtrlButtonBack;
	CButton m_CtrlButtonForward;
	CButton m_CtrlButtonReload;
	CButton m_CtrlButtonReCache;
	CEdit m_CtrlEditFind;
	CButton m_CtrlButtonFind;
	CMFCTabCtrl m_CtrlTabViewer;
	afx_msg void OnBnClickedButtonLoad();
	afx_msg void OnBnClickedButtonBack();
	afx_msg void OnBnClickedButtonForward();
	afx_msg void OnBnClickedButtonReload();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg void OnDestroy();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnBnClickedButtonRecache();
	afx_msg void OnBnClickedButtonFind();

public:
	CString m_WebViewTempFolder;	// WebView2の一時フォルダの場所
	wil::com_ptr<ICoreWebView2Environment> m_pWebEnvironment;	// environmentはアプリケーション内で使いまわす
	CRect m_DefParentRect;	// ウィンドウの最小サイズ
	int m_TabViewWndId;	// nex ID of created window; never decreases

	inja::json m_Issues;	// issues data

	// common JSON data for issue
	inja::Template m_IssueTemplate;	// holds Issue.html data
	inja::Environment m_InjaEnv;	// holds callback functions
	std::map<int, std::string> m_Members;
	std::map<int, std::string> m_Statuses;
	std::map<int, std::string> m_Trackers;
	std::map<int, std::string> m_Priorities;

	nlohmann::json ReadJson(const wchar_t* filePath);
	void AddTab();
	void AddTab(CString file);

private:
	void LoadCommonData();
	void LoadSetting();
	void SaveSetting();
	void LoadIssues();
	void SetupCallbacks();

	void SearchIssueList(std::map<int, std::vector<std::string>>& result, const nlohmann::json& issueList, const std::string& query);
	void SearchIssue(std::vector<std::string>& result, const nlohmann::json& issue, const std::string& query, const std::string& path = "");
	virtual void PostNcDestroy();
};

// inja callbacks
/// <summary>Check if 1st argument is a string</summary>
bool CallbackIs1stArgString(const inja::Arguments& args, inja::json& errText);
/// <summary>Convert UTC time string to local time string</summary>
inja::json CallbackUtcToLocal(inja::Arguments& args);
/// <summary> Convert UTC time string to "x ago" format string</summary>
inja::json CallbackUtcToAgo(inja::Arguments& args);
/// <summary>Convert UTC time string to "YYYY/MM/DD" format string</summary>
inja::json CallbackUtcToYMD(inja::Arguments& args);
