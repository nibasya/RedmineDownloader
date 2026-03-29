#pragma once
#include <afxwin.h>

class CRedmineViewerDlg;

class CIssueWnd :
    public CWnd
{
public:
	DECLARE_MESSAGE_MAP()

public:
	CRedmineViewerDlg* m_pParent;

	wil::com_ptr<ICoreWebView2Controller> m_WebViewController;
	wil::com_ptr<ICoreWebView2> m_WebView;

	CString m_IssueFilePath;	// issue file path; used for browsing
	CString m_RequestedUrl;	// last requested url; might not be used for browsing

	CIssueWnd(CRedmineViewerDlg *pParent);
	~CIssueWnd();
	bool CreateWebView2(UINT nID, LPCWSTR url);

	HRESULT WebView2CreateController(HRESULT result, ICoreWebView2Controller* controller);
	bool WebView2GetLocalFilePathFromUri(const wil::unique_cotaskmem_string& uri, CString& outPath);
	HRESULT NewWindowRequestHandler(ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args);	// 新規ウィンドウのリクエストを完全にブロックしつつuriを取得するイベントハンドラー
	HRESULT NavigationStartingHandler(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args);	// リンククリック時のイベントハンドラー（ローカルファイルへのリンクを既定のアプリで開くため）
	void ShowMessageBox(const CString& msg, const CString& title = CString(L""));

	bool ShowIssue();
	CString SetTabTitle(CString filePath);
	void ReplaceId(nlohmann::json& json, std::string property, std::string name, std::map<int, std::string> data);
	std::string ReplaceIssueLink(std::string text);
	std::string ConvertMdToHtml(std::string mdText);		// Converts Markdown text into HTML body
	static void ConvertMdToHtmlSub(const MD_CHAR* text, MD_SIZE size, void* userData);	// callback for ConvertMDToHtml
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnClose();
	afx_msg void OnDestroy();
	virtual void PostNcDestroy();
};

