// pch.h: プリコンパイル済みヘッダー ファイルです。  
// 次のファイルは、その後のビルドのビルド パフォーマンスを向上させるため 1 回だけコンパイルされます。  
// コード補完や多くのコード参照機能などの IntelliSense パフォーマンスにも影響します。  
// ただし、ここに一覧表示されているファイルは、ビルド間でいずれかが更新されると、すべてが再コンパイルされます。  
// 頻繁に更新するファイルをここに追加しないでください。追加すると、パフォーマンス上の利点がなくなります。  

#ifndef PCH_H  
#define PCH_H  

// プリコンパイルするヘッダーをここに追加します  
#include "framework.h"  

#if defined(_MSC_VER) && (_MSC_VER < 1950)
	#pragma warning(push)
	#pragma warning(disable:4996)
	#include <cpprest/json.h>
	#include <cpprest/filestream.h>
	#include <cpprest/http_client.h>
	#pragma warning(pop)
#endif // _MSC_VER < 1950

#endif //PCH_H
