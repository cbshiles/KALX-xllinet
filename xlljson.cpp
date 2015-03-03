// xlljson.cpp - JSON objects in Excel
#define EXCEL12
#include "xll/xll.h"
#include "xlljson.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xstring xstring;
typedef traits<XLOPERX>::xword xword;

static AddInX xai_json_set(
	FunctionX(XLL_HANDLEX, _T("?xll_json_set"), _T("JSON.SET"))
	.Arg(XLL_LPOPERX, _T("Object"), _T("is a range containing the JSON object"))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a handle to a JSON string object"))
	.Documentation()
);
HANDLEX WINAPI xll_json_set(LPOPERX po)
{
#pragma XLLEXPORT
	handlex h;

	try {
		h = handle<xstring>(new xstring(json::stringify(*po))).get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}


static AddInX xai_json_get(
	FunctionX(XLL_LPOPERX, _T("?xll_json_get"), _T("JSON.GET"))
	.Arg(XLL_LPOPERX, _T("Handle"), _T("is a handle to a JSON object"))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a JSON object"))
	.Documentation()
);
LPOPERX WINAPI xll_json_get(LPOPERX po)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		if (po->xltype == xltypeNum) {
			handle<xstring> hs(po->val.num);
			xcstr cs = hs->c_str();
			o = json::parse_object<XLOPERX>(&cs);
		}
		else if (po->xltype == xltypeStr) {
			xstring s(po->val.str + 1, po->val.str + 1 + po->val.str[0]);
			xcstr cs = s.c_str();
			o = json::parse_object<XLOPERX>(&cs);
		}
		else {
			XLL_WARNING("JSON.GET: argument must be a string or a handle to a string");
			o = OPERX(xlerr::NA);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}


static AddInX xai_range_expand(_T("?xll_range_expand"), _T("XLL.EXPAND"));
int WINAPI xll_range_expand(void)
{
#pragma XLLEXPORT
	try {
		OPERX ac = XLL_XLF(ActiveCell);
		OPERX co = XLL_XL_(Coerce, ac);
		ensure (co.xltype == xltypeNum);
		handle<OPERX> ho(co.val.num);
		OPERX off = XLL_XLF(Offset, ac, OPERX(0), OPERX(1), OPERX(ho->rows()), OPERX(ho->columns()));
		XLL_XL_(Set, off, *ho);
		XLL_XLC(Select, off);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

int
xll_expand_close(void)
{
	try {
		if (Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Expand")))
			Excel<XLOPER>(xlfDeleteCommand, OPER(7), OPER(4), OPER("Expand"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		
		return 0;
	}

	return 1;
}
static Auto<Close> xac_expand(xll_expand_close);

int
xll_expand_open(void)
{
	try {
		// Try to add just above first menu separator.
		OPER oPos;
		oPos = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("-"));
		oPos = 6;

		OPER oAdj = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Expand"));
		if (oAdj == OPER(xlerr::NA)) {
			OPER oAdj(1, 5);
			oAdj(0, 0) = "Expand";
			oAdj(0, 1) = "XLL.EXPAND";
			oAdj(0, 3) = "Expand output of array function.";
			Excel<XLOPER>(xlfAddCommand, OPER(7), OPER(4), oAdj, oPos);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_expand(xll_expand_open);
