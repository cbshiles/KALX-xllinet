// xllquandle.cpp
#include "xll/xll.h"

#define CATEGORY _T("Inet")
#define QUANDLE_KEY _T("ZDZV_AyuBBJR6f_hNkrS")

using namespace xll;

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;

static AddInX xai_quandle(
	FunctionX(XLL_CSTRINGX, _T("?xll_quandle"), _T("QUANDLE"))
	.Arg(XLL_CSTRINGX, _T("Database"), _T("is the quandle database code"))
	.Arg(XLL_CSTRINGX, _T("Table"), _T("is a quandle table code"))
	.Arg(XLL_CSTRINGX, _T("Query"), _T("is an optional query string."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a Quandle query"))
	.Documentation()
);
xcstr WINAPI xll_quandle(xcstr db, xcstr tb, xcstr qu)
{
#pragma XLLEXPORT
	static xstring q;

	try {
		q = _T("https://www.quandl.com/api/v1/datasets");

		if (db)
			q.append(_T("/")).append(db);
		if (tb)
			q.append(_T("/")).append(tb);
		q.append(_T(".json?auth_token="));
		q.append(QUANDLE_KEY);
		q.append(qu);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return q.c_str();
}

static AddInX xai_quandl_metadata(
	FunctionX(XLL_LPOPERX, _T("?xll_quandle_metadata"), _T("QUANDLE.METADATA"))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a string indicating the metadata for the table should be returned."))
	.Documentation()
);
LPOPERX WINAPI xll_quandle_metadata(void)
{
#pragma XLLEXPORT
	static OPERX md(_T("?exclude_data=true"));

	return &md;
}

static AddInX xai_quandl_ascending(
	FunctionX(XLL_LPOPERX, _T("?xll_quandle_ascending"), _T("QUANDLE.ASCENDING"))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a string indicating dates should be in ascending order. Default is descending."))
	.Documentation()
);
LPOPERX WINAPI xll_quandle_ascending(void)
{
#pragma XLLEXPORT
	static OPERX asc(_T("?sort_order=asc"));

	return &asc;
}

static AddInX xai_quandl_rows(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_rows"), _T("QUANDLE.ROWS"))
	.Arg(XLL_WORDX, _T("Rows"), _T("is the number of rows to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Specify the number of rows to be returned."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_rows(WORD rows)
{
#pragma XLLEXPORT
	static OPERX rows_(_T("?rows="));

	return XLL_XLF(Concatenate, rows_, OPERX(rows)).XLFree();
}

static AddInX xai_quandl_column(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_column"), _T("QUANDLE.COLUMN"))
	.Arg(XLL_WORDX, _T("Column"), _T("is the column to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Specify the column to be returned. The first column is 0 and contains the dates."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_column(WORD column)
{
#pragma XLLEXPORT
	static OPERX column_(_T("?column="));

	return XLL_XLF(Concatenate, column_, OPERX(column)).XLFree();
}

static AddInX xai_quandl_start(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_start"), _T("QUANDLE.START"))
	.Arg(XLL_DOUBLEX, _T("Start"), _T("is starting date of data to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the starting date of the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_start(double start)
{
#pragma XLLEXPORT
	static OPERX start_(_T("?trim_start=")), format(_T("yyyy-mm-dd"));

	return XLL_XLF(Concatenate, start_, XLL_XLF(Text, OPERX(start), format)).XLFree();
}

static AddInX xai_quandl_end(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_end"), _T("QUANDLE.END"))
	.Arg(XLL_DOUBLEX, _T("end"), _T("is ending date of data to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the ending date of the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_end(double end)
{
#pragma XLLEXPORT
	static OPERX end_(_T("?trim_end=")), format(_T("yyyy-mm-dd"));

	return XLL_XLF(Concatenate, end_, XLL_XLF(Text, OPERX(end), format)).XLFree();
}

static AddInX xai_quandl_frequency(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_frequency"), _T("QUANDLE.FREQUENCY"))
	.Arg(XLL_CSTRINGX, _T("Freq"), _T("one of none,daily,weekly,monthly,quaterly,annual."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the frequencying the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_frequency(xcstr freq)
{
#pragma XLLEXPORT
	static OPERX collapse(_T("?collapse="));

	OPERX f;

	switch (::towlower(*freq)) {
	case 'd':
		f = _T("daily");
		break;
	case 'w':
		f = _T("weekly");
		break;
	case 'm':
		f = _T("monthly");
		break;
	case 'q':
		f = _T("quarterly");
		break;
	case 'a':
		f = _T("annual");
		break;
	default:
		f = _T("none");
	}

	return XLL_XLF(Concatenate, collapse, f).XLFree();
}

static AddInX xai_quandl_calculate(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_calculate"), _T("QUANDLE.CALCULATE"))
	.Arg(XLL_CSTRINGX, _T("Transformation"), _T("one of diff, rdiff, cumul, or normalize."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the calculateing the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_calculate(xcstr trans)
{
#pragma XLLEXPORT
	static OPERX collapse(_T("?transformation="));

	OPERX f;

	switch (::towlower(*trans)) {
	case 'd':
		f = _T("diff");
		break;
	case 'r':
		f = _T("rdiff");
		break;
	case 'c':
		f = _T("cumul");
		break;
	case 'n':
		f = _T("normalize");
		break;
	default:
		f = _T("none");
	}

	return XLL_XLF(Concatenate, trans, f).XLFree();
}

static AddInX xai_quandl_search(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandle_search"), _T("QUANDLE.SEARCH"))
	.Arg(XLL_DOUBLEX, _T("Terms"), _T("one or more words to seach for."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the searching the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandle_search(LPOPERX term)
{
#pragma XLLEXPORT
	static OPERX query, space(_T(" ")), plus(_T("+"));
	
	try {
		query = _T("?query=");
		for (const auto& t : query) {
			query = XLL_XLF(Concatenate, query, plus, XLL_XLF(Substitute, t, space, plus));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		query = OPERX(xlerr::NA);
	}


	return &query;
}