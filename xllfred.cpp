// xllfred.cpp - http://api.stlouisfed.org/docs/fred/
#define EXCEL12
#include "xll/xll.h"
#include "xll/utility/registry.h"

#define CATEGORY _T("FRED")
#define FRED_URL _T("http://api.stlouisfed.org/fred")
#define FRED_KEY _T("1947fc010512c76edbe4b8f6f8e0bdd5")

using namespace xll;

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;

static std::basic_string<TCHAR> fred_key;
int xll_get_fred_key(void)
{
	try {
		Reg::Object<TCHAR, std::basic_string<TCHAR>> FredKey(HKEY_CURRENT_USER, _T("Software\\KALX\\Inet"), _T("FredlKey"));
		if (FredKey == _T("")) {
			OPERX key = XLL_XLF(Input, OPERX(_T("Enter your API key from http://api.stlouisfed.org")), OPERX(2), OPERX(_T("Enter API Key")));
			ensure (key.xltype == xltypeStr);
			FredKey = key.string();
		}
		fred_key = FredKey;
	}
	catch (...) {

		return FALSE;
	}

	return TRUE;
}
static Auto<OpenX> xao_get_fred_key(xll_get_fred_key);

static AddInX xai_fred(
	FunctionX(XLL_CSTRINGX, _T("?xll_fred"), _T("FRED"))
	.Arg(XLL_CSTRINGX, _T("Topic"), _T("is a topic string."))
	.Arg(XLL_CSTRINGX, _T("_Query"), _T("is an optional query string."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a FRED URL."))
	.Documentation()
);
xcstr WINAPI xll_fred(xcstr topic, xcstr qu)
{
#pragma XLLEXPORT
	static xstring q;

	try {
		q = FRED_URL;
		if (topic && *topic) {
			q.append(_T("/"));
			q.append(topic);
		}

		q.append(_T("?file_type=json&api_key="));
		q.append(fred_key);
		if (qu && *qu)
			q.append(qu);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return q.c_str();
}

static AddInX xai_fred_category(
	FunctionX(XLL_CSTRINGX, _T("?xll_fred_category"), _T("FRED.CATEGORY"))
	.Arg(XLL_WORDX, _T("Id"), _T("is the FRED category Id. Default is 0, the root category."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a FRED URL for the category Id."))
	.Documentation()
);
xcstr WINAPI xll_fred_category(WORD id)
{
#pragma XLLEXPORT
	OPERX q;

	try {
		q = XLL_XLF(Concatenate, OPERX(_T("&category_id=")), OPERX(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return xll_fred(_T("category"), q.string().c_str());
}

static AddInX xai_fred_category_children(
	FunctionX(XLL_CSTRINGX, _T("?xll_fred_category_children"), _T("FRED.CATEGORY.CHILDREN"))
	.Arg(XLL_WORDX, _T("Id"), _T("get the child categories for a specified parent category. Default is 0, the root category."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a FRED URL for the category_children Id."))
	.Documentation()
);
xcstr WINAPI xll_fred_category_children(WORD id)
{
#pragma XLLEXPORT
	OPERX q;

	try {
		q = XLL_XLF(Concatenate, OPERX(_T("&category_id=")), OPERX(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return xll_fred(_T("category/children"), q.string().c_str());
}

static AddInX xai_fred_category_related(
	FunctionX(XLL_CSTRINGX, _T("?xll_fred_category_related"), _T("FRED.CATEGORY.RELATED"))
	.Arg(XLL_WORDX, _T("Id"), _T("get the series for a specified category"))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a FRED URL for the category_related Id."))
	.Documentation()
);
xcstr WINAPI xll_fred_category_related(WORD id)
{
#pragma XLLEXPORT
	OPERX q;

	try {
		q = XLL_XLF(Concatenate, OPERX(_T("&category_id=")), OPERX(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return xll_fred(_T("category/related"), q.string().c_str());
}

static AddInX xai_fred_category_series(
	FunctionX(XLL_CSTRINGX, _T("?xll_fred_category_series"), _T("FRED.CATEGORY.SERIES"))
	.Arg(XLL_WORDX, _T("Id"), _T("is the series id."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a FRED URL for the category_series Id."))
	.Documentation()
);
xcstr WINAPI xll_fred_category_series(WORD id)
{
#pragma XLLEXPORT
	OPERX q;

	try {
		q = XLL_XLF(Concatenate, OPERX(_T("&category_id=")), OPERX(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return xll_fred(_T("category/series"), q.string().c_str());
}
