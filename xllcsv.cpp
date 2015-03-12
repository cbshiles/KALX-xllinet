// xllcsv.cpp - comma separated value functions
#include "xllcsv.h"

using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xstring xstring;

static AddInX xai_csv_get(
	FunctionX(XLL_LPOPERX, _T("?xll_csv_get"), _T("CSV.GET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a std::string object"))
	.Arg(XLL_CSTRINGX, _T("_FP"), _T("is the optional field separator. Default is \",\"."))
	.Arg(XLL_CSTRINGX, _T("_RS"), _T("is the optional record separator. Default is \";|\\n\"."))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Parse comma separated values based on field separator and record separator."))
	.Documentation()
);
LPOPERX WINAPI xll_csv_get(HANDLEX h, xcstr fs, xcstr rs)
{
#pragma XLLEXPORT
	static OPERX t;

	try {
		t.resize(0,0);

		if (!*fs)
			fs = _T(",");
		if (!*rs)
			rs = _T(";\n");

		handle<xstring> hs(h);
		xcstr cs = hs->c_str();
		t = csv::parse(cs, fs, rs);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		t = OPERX(xlerr::NA);
	}

	return &t;
}

static AddInX xai_csv_set(
	FunctionX(XLL_HANDLEX, _T("?xll_csv_set"), _T("CSV.SET"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a range"))
	.Arg(XLL_CSTRINGX, _T("_FP"), _T("is the optional field separator. Default is \",\"."))
	.Arg(XLL_CSTRINGX, _T("_RS"), _T("is the optional record separator. Default is \"\\n\"."))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a handle to a string of comma separated values."))
	.Documentation()
);
HANDLEX WINAPI xll_csv_set(LPOPERX po, xcstr fs, xcstr rs)
{
#pragma XLLEXPORT
	handlex h;

	try {
		if (!*fs)
			fs = _T(",");
		if (!*rs)
			rs = _T("\n");

		handle<xstring> hs{new xstring(csv::to_string(*po, fs, rs))};

		h = hs.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

#ifdef _DEBUG

inline void test_csv_trim(xcstr s, xchar l, xchar r)
{
	xcstr b = s, e = s + _tcslen(s);
	csv::trim(&b, &e);
	ensure(b == s + l);
	ensure(e == s + _tcslen(s) - r);
}
	
void xll_test_csv_trim()
{
	test_csv_trim(_T(""), 0, 0);
	test_csv_trim(_T(" "), 1, 0);
	test_csv_trim(_T(" a"), 1, 0);
	test_csv_trim(_T("a "), 0, 1);
	test_csv_trim(_T(" \ta"), 2, 0);
	test_csv_trim(_T(" a\t\r\n "), 1, 4);
}

void xll_test_csv_unquote()
{
	xcstr s = _T("\"hi\"");
	xcstr bs = s, es = s + _tcslen(s);
	csv::unquote(&bs, &es);
	ensure(bs == s + 1);
	ensure(es == s + _tcslen(s) - 1);

}

int xll_test_csv()
{
	try {
		xll_test_csv_trim();
		xll_test_csv_unquote();
			
		OPERX o {
			OPERX(xlerr::Div0)};//, OPERX(_T("abc")),
//			OPERX(false), OPERX(xlerr::NA)
//		};
//		o.resize(2,2);
		xstring csv = csv::to_string(o);
//		OPERX o_;
//		o_ = csv::parse(csv.c_str());
		OPERX p_, q_;
		OPERX o_(_T("foo"), 3);
		o.resize(1,1);
		p_.push_back(o_);
		p_.resize(p_.rows(), p_.columns());
		q_.push_back(p_);
//		q_.push_back(p_.resize(p_.rows(), p_.columns()));

//		ensure (o_ == o);

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static Auto<OpenAfterX> xao_test_csv(xll_test_csv);

#endif // _DEBUG