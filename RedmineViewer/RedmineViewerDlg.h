
// RedmineViewerDlg.h : ヘッダー ファイル
//

#pragma once

#include <wil/com.h>
#include "WebView2.h"

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

	wil::com_ptr<ICoreWebView2Controller> m_WebViewController;
	wil::com_ptr<ICoreWebView2> m_WebView;

public:
	CButton m_CtrlButtonLoad;
	CButton m_CtrlButtonBack;
	CButton m_CtrlButtonForward;
	CStatic m_CtrlWebView;
	afx_msg void OnBnClickedButtonLoad();
	afx_msg void OnBnClickedButtonBack();
	afx_msg void OnBnClickedButtonForward();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg void OnDestroy();
	afx_msg void OnDropFiles(HDROP hDropInfo);
private:
	bool LoadJson(const wchar_t* filePath);
};
