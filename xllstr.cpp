// xllstr.cpp - std::basic_string wrappers
#define EXCEL12
#include "xll/xll.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xstring xstring;

static AddInX xai_str_set(
	FunctionX(XLL_HANDLEX, _T("?xll_str_set"), _T("STR.SET"))
	.Arg(XLL_PSTRINGX, _T("String"), _T("is a string of characters."))
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
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a string."))
	.Documentation()
);
xcstr WINAPI xll_str_get(HANDLEX h)
{
#pragma XLLEXPORT
	xcstr s{0};

	try {
		handle<xstring> hs(h);
//		ensure (hs->length() <= traits<XLOPERX>::strmax);
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
