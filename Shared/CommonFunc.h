#pragma once


CString GetFileNameFromJson(const CString& fileName, const CString& ID) {
	CString ret;
	ret.Format(L"%s_%s", (LPCWSTR)ID, (LPCWSTR)fileName);
	return ret;
};