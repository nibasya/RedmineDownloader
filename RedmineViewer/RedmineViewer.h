
// RedmineViewer.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH に対してこのファイルをインクルードする前に 'pch.h' をインクルードしてください"
#endif

#include "resource.h"		// メイン シンボル


// CRedmineViewerApp:
// このクラスの実装については、RedmineViewer.cpp を参照してください
//

class CRedmineViewerApp : public CWinApp
{
public:
	CRedmineViewerApp();

// オーバーライド
public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()

public:
	CString m_AppFolderPath;	// アプリケーションのフォルダパス
};

extern CRedmineViewerApp theApp;
