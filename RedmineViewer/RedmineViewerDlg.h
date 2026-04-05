
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
	afx_msg LPARAM OnShowMessageBox(WPARAM wParam, LPARAM lParam);
	virtual void PostNcDestroy();

public:
	CString m_AppFolderPath;	// アプリケーションのフォルダパス
	CString m_WebViewTempFolder;	// WebView2の一時フォルダの場所
	wil::com_ptr<ICoreWebView2Environment> m_pWebEnvironment;	// environmentはアプリケーション内で使いまわす
	CRect m_DefParentRect;	// ウィンドウの最小サイズ
	int m_TabViewWndId;	// nex ID of created window; never decreases

	inja::json m_Issues;	// issues data

	// common JSON data for issue
	inja::Template m_IssueTemplate;	// holds Issue.html data
	inja::Template m_SearchResultTemplate;	// holds Search result.html data
	inja::Template m_ProjectTemplate;	// holds ProjectInfo.html data
	inja::Environment m_InjaEnv;	// holds callback functions
	std::map<int, std::string> m_Members;
	std::map<int, std::string> m_Statuses;
	std::map<int, std::string> m_Trackers;
	std::map<int, std::string> m_Priorities;

	/// <summary>
	/// wrapper function of json::pharse() with file check and error message box
	/// </summary>
	/// <param name="filePath">path with filename of the JSON file</param>
	/// <param name ="suppressMessageBox">if true, suppress message box when error occurs and throw std::exception</param>
	/// <returns>Returns read json. If any error, show error message box and return empty json.</returns>
	nlohmann::json ReadJson(const wchar_t* filePath, bool suppressMessageBox = false);
	void AddStartTab();
	void AddTab(CString file);

private:
	void LoadCommonData();
	void LoadSetting();
	void SaveSetting();
	void LoadIssues(bool reflesh);
	nlohmann::json ScanLocalFolder();	// scan local files to find issues and list them up for UpdateCache()
	nlohmann::json CreateUpdateList(nlohmann::json& downloadedList);
	void UpdateCache(const nlohmann::json& updateList);
	void SetupCallbacks();

	void SearchIssueList(std::vector<nlohmann::json>& result, const nlohmann::json& issueList, const std::string& query);
	void SearchIssue(nlohmann::json& result, const nlohmann::json& issue, const nlohmann::json& givenKey, const std::string& query);
	/// <summary>
	/// Extracts text including key from str with length limit
	/// </summary>
	/// <param name="str">source text</param>
	/// <param name="key">key</param>
	/// <param name="length">total length of extracted text</param>
	/// <returns>extracted text</returns>
	CString ExtractContext(const CString& str, const CString& key, int length);
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
/// <summary>Limit the length of the string</summary>
/// <param name="args">1st argument: string to limit, 2nd argument: length limit</param>
inja::json CallbackLimitStr(inja::Arguments& args);

// user-defined messeges
#define WM_USER_SHOW_MESSAGE_BOX (WM_USER + 101)