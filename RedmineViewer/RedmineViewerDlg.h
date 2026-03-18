
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
	wil::com_ptr<ICoreWebView2Controller> m_WebViewController;
	wil::com_ptr<ICoreWebView2> m_WebView;
	CString m_WebViewTempFolder;

	HRESULT WebView2CreateController(HRESULT result, ICoreWebView2Controller* controller);
	HRESULT NewWindowRequestHandler(ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args);	// 新規ウィンドウのリクエストを完全にブロックしつつuriを取得するイベントハンドラー

public:
	CButton m_CtrlButtonLoad;
	CButton m_CtrlButtonBack;
	CButton m_CtrlButtonForward;
	CStatic m_CtrlWebView;
	CButton m_CtrlButtonReload;
	afx_msg void OnBnClickedButtonLoad();
	afx_msg void OnBnClickedButtonBack();
	afx_msg void OnBnClickedButtonForward();
	afx_msg void OnBnClickedButtonReload();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg void OnDestroy();
	afx_msg void OnDropFiles(HDROP hDropInfo);

private:
	CString m_IssueFilePath;
	inja::Template m_IssueTemplate;
	inja::Environment m_Env;
	std::map<int, std::string> m_Members;
	std::map<int, std::string> m_Statuses;
	std::map<int, std::string> m_Trackers;
	std::map<int, std::string> m_Priorities;

	void LoadCommonData();
	nlohmann::json ReadJson(const wchar_t* filePath);
	bool ShowIssue();
	void ReplaceId(nlohmann::json& json, std::string property, std::string name, std::map<int, std::string> data);
	std::string ConvertMdToHtml(std::string mdText);		// Converts Markdown text into HTML body
	static void ConvertMdToHtmlSub(const MD_CHAR* text, MD_SIZE size, void* userData);	// callback for ConvertMDToHtml
	void LoadSetting();
	void SaveSetting();
	void SetupCallbacks();
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
