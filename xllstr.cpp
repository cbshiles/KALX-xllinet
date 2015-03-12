// xllstr.cpp - std::basic_string wrappers
#define EXCEL12
#include "xll/xll.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xstring xstring;

static AddInX xai_str_set(
	FunctionX(XLL_HANDLEX, _T("?xll_str_set"), _T("STR.SET"))
	.Arg(XLL_PSTRINGX, _T("String"), _T("is a string of characters."))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a handle to a string."))
	.Documentation()
);
HANDLEX WINAPI xll_str_set(xcstr s)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<xstring> hs(new xstring(s + 1, s[0]));

		h = hs.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddInX xai_str_get(
	FunctionX(XLL_CSTRINGX, _T("?xll_str_get"), _T("STR.GET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a std::string object"))
	.Arg(XLL_WORDX, _T("_Pos"), _T("is the initial position of a substring to return. Default is 0."))
	.Arg(XLL_WORDX, _T("_Len"), _T("is an optional length of the substring to return. Default is the entire string."))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a string or substring."))
	.Documentation()
);
xcstr WINAPI xll_str_get(HANDLEX h, WORD pos, WORD len)
{
#pragma XLLEXPORT
	static xstring str;
	xcstr s{0};

	try {
		handle<xstring> hs(h);

		if (len > 0) {
			str = hs->substr(pos, len);
			s = str.c_str();
		}
		else
			s = hs->c_str();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return s;
}

static AddInX xai_str_len(
	FunctionX(XLL_LONGX, _T("?xll_str_len"), _T("STR.LEN"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a std::string object"))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return the length of a string"))
	.Documentation()
);
LONG WINAPI xll_str_len(HANDLEX h)
{
#pragma XLLEXPORT
	LONG len{-1};

	try {
		handle<xstring> hs(h);
		
		len = hs->length();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return len;
}

static AddInX xai_str_append(
	FunctionX(XLL_HANDLEX, _T("?xll_str_append"), _T("STR.APPEND"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a std::string object"))
	.Arg(XLL_CSTRINGX, _T("Str"), _T("is a string to append."))
	.Arg(XLL_WORDX, _T("_Len"), _T("is an optional number of characters to copy from Str. Default is all."))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Append Str to the string associated with Handle."))
	.Documentation()
);
HANDLEX WINAPI xll_str_append(HANDLEX h, xcstr str, WORD n)
{
#pragma XLLEXPORT
	try {
		handle<xstring> hs(h);
		
		if (n)
			hs->append(str, n);
		else
			hs->append(str);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}
