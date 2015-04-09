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
	.Arg(XLL_LONGX, _T("_Pos"), _T("is the initial position of a substring to return. Default is 0."))
	.Arg(XLL_LONGX, _T("_Len"), _T("is an optional length of the substring to return. Default is the entire string."))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a string or substring. If Pos is negative characters are taken from the end of the string."))
	.Documentation()
);
xcstr WINAPI xll_str_get(HANDLEX h, LONG pos, LONG len)
{
#pragma XLLEXPORT
	static xstring str;
	xcstr s{0};

	try {
		handle<xstring> hs(h);

		if (pos > 0) {
			if (len == 0)
				len = static_cast<LONG>(hs->length()) - pos;
			str = hs->substr(pos, len);
			s = str.c_str();
		}
		if (pos < 0) {
			size_t pos_ = hs->length() + pos;
			if (len == 0)
				len = -pos;
			str = hs->substr(pos_, len);
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
		
		len = static_cast<LONG>(hs->length());
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

static AddInX xai_str_split(
	FunctionX(XLL_LPOPERX, _T("?xll_str_split"), _T("STR.SPLIT"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a std::string object"))
	.Arg(XLL_CSTRINGX, _T("_Delim"), _T("is an optional string of delimiters. Default is newline."))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Split the string associated with Handle."))
	.Documentation()
);
LPOPERX WINAPI xll_str_split(HANDLEX h, xcstr delim)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		handle<xstring> hs(h);

		static xcstr nl = _T("\n");
		xcstr d = *delim ? delim : nl;

		using st = xstring::size_type;
		st b = hs->find_first_not_of(d);
		while (b != xstring::npos) {
			st e = hs->find_first_of(d, b);
			o.push_back(hs->substr(b, e - b));
			b = hs->find_first_not_of(d, b + e);
		}

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

#ifdef _DEBUG

XLL_TEST_BEGIN(test_str)

	handle<xstring> hs = new xstring{_T("a;b;c")};
	OPERX o = *xll_str_split(hs.get(), _T(";"));

	ensure (o.size() == 3);
	ensure (o[0] == _T("a"));
	ensure (o[1] == _T("b"));
	ensure (o[2] == _T("c")); // fails!!!

XLL_TEST_END(test_str)

#endif // _DEBUG