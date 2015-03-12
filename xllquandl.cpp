// xllquandl.cpp
#define EXCEL12
#include "xll/xll.h"
#include "xll/utility/registry.h"

#define CATEGORY _T("Inet")
#define QUANDL_KEY _T("ZDZV_AyuBBJR6f_hNkrS")

using namespace xll;

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;

static std::basic_string<TCHAR> quandl_key;
int xll_get_quandl_key(void)
{
	try {
		Reg::Object<TCHAR, std::basic_string<TCHAR>> Quandl_key(HKEY_CURRENT_USER, _T("KALX\\QUANDL"), _T("Key"), _T(""));
		if (Quandl_key == _T("")) {
			OPERX key = XLL_XLF(Input, OPERX(_T("Enter your API key from http://quandl.com")), OPERX(2), OPERX(_T("Enter API Key")));
			quandl_key = key.to_string();
		}
		else {
			quandl_key = Quandl_key;
		}
	}
	catch (...) {
		quandl_key = QUANDL_KEY;

		return FALSE;
	}

	return TRUE;
}
static Auto<OpenX> xao_get_quandl_key(xll_get_quandl_key);

static AddInX xai_quandl(
	FunctionX(XLL_CSTRINGX, _T("?xll_quandl"), _T("QUANDL"))
	.Arg(XLL_CSTRINGX, _T("Database"), _T("is the Quandl database code"))
	.Arg(XLL_CSTRINGX, _T("Table"), _T("is a Quandl table code"))
	.Arg(XLL_CSTRINGX, _T("Query"), _T("is an optional query string."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a Quandl query"))
	.Documentation()
);
xcstr WINAPI xll_quandl(xcstr db, xcstr tb, xcstr qu)
{
#pragma XLLEXPORT
	static xstring q;

	try {
		q = _T("https://www.quandl.com/api/v1/datasets");

		if (db && *db)
			q.append(_T("/")).append(db);
		if (tb && *tb)
			q.append(_T("/")).append(tb);
		q.append(_T(".json?auth_token="));
		q.append(quandl_key);
		while (q.back() == 0)
			q.pop_back(); // ???
		q.append(qu);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return q.c_str();
}

static AddInX xai_quandl_metadata(
	FunctionX(XLL_LPOPERX, _T("?xll_quandl_metadata"), _T("QUANDL.METADATA"))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a string indicating the metadata for the table should be returned."))
	.Documentation()
);
LPOPERX WINAPI xll_quandl_metadata(void)
{
#pragma XLLEXPORT
	static OPERX md(_T("&exclude_data=true"));

	return &md;
}

static AddInX xai_quandl_query(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_query"), _T("QUANDL.QUERY"))
	.Arg(XLL_LPOPERX, _T("Terms"), _T("is a range containg terms to query."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a string indicating the metadata for the table should be returned."))
	.Documentation()
	);
LPXLOPERX WINAPI xll_quandl_query(LPOPERX pq)
{
#pragma XLLEXPORT
	static OPERX q(_T("&query="));

	OPERX Q((*pq)[0]);

	for (xword i = 1; i < pq->size(); ++i) {
		if ((*pq)[i]) {
			Q.append(_T("+"), 1);
			Q.append((*pq)[i]);
		}
	}

	return XLL_XLF(Concatenate, q, Q).XLFree();
}

static AddInX xai_quandl_ascending(
	FunctionX(XLL_LPOPERX, _T("?xll_quandl_ascending"), _T("QUANDL.ASCENDING"))
	.Arg(XLL_BOOLX, _T("_Asc"), _T("is an optional boolean indicating the dates should be in ascending order. Default is FALSE."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a string indicating dates should be in ascending order. Default is descending."))
	.Documentation()
);
LPOPERX WINAPI xll_quandl_ascending(BOOL asc)
{
#pragma XLLEXPORT
	static OPERX Asc(_T("&sort_order=asc")), Des(_T(""));

	return asc ? &Asc : &Des;
}

static AddInX xai_quandl_rows(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_rows"), _T("QUANDL.ROWS"))
	.Arg(XLL_WORDX, _T("Rows"), _T("is the number of rows to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Specify the number of rows to be returned."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandl_rows(WORD rows)
{
#pragma XLLEXPORT
	static OPERX rows_(_T("&rows="));

	return XLL_XLF(Concatenate, rows_, OPERX(rows)).XLFree();
}

static AddInX xai_quandl_column(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_column"), _T("QUANDL.COLUMN"))
	.Arg(XLL_WORDX, _T("Column"), _T("is the column to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Specify the column to be returned. The first column is 0 and contains the dates."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandl_column(WORD column)
{
#pragma XLLEXPORT
	static OPERX column_(_T("&column="));

	return XLL_XLF(Concatenate, column_, OPERX(column)).XLFree();
}

static AddInX xai_quandl_start(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_start"), _T("QUANDL.START"))
	.Arg(XLL_DOUBLEX, _T("Start"), _T("is starting date of data to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the starting date of the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandl_start(double start)
{
#pragma XLLEXPORT
	static OPERX start_(_T("&trim_start=")), format(_T("yyyy-mm-dd"));

	return XLL_XLF(Concatenate, start_, XLL_XLF(Text, OPERX(start), format)).XLFree();
}

static AddInX xai_quandl_end(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_end"), _T("QUANDL.END"))
	.Arg(XLL_DOUBLEX, _T("end"), _T("is ending date of data to return."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the ending date of the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandl_end(double end)
{
#pragma XLLEXPORT
	static OPERX end_(_T("&trim_end=")), format(_T("yyyy-mm-dd"));

	return XLL_XLF(Concatenate, end_, XLL_XLF(Text, OPERX(end), format)).XLFree();
}

static AddInX xai_quandl_frequency(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_frequency"), _T("QUANDL.FREQUENCY"))
	.Arg(XLL_CSTRINGX, _T("Freq"), _T("one of none,daily,weekly,monthly,quaterly,annual."))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the frequencying the data."))
	.Documentation()
);
LPOPERX WINAPI xll_quandl_frequency(xcstr freq)
{
#pragma XLLEXPORT
	static OPERX c(_T("&collapse=")), f;

	try {
		f = _T("");

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
		}

		if (f != _T(""))
			f = XLL_XLF(Concatenate, c, f);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		f = OPERX(xlerr::NA);
	}

	return &f;
}

static AddInX xai_quandl_calculate(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_calculate"), _T("QUANDL.CALCULATE"))
	.Arg(XLL_CSTRINGX, _T("Transformation"), _T("one of diff, rdiff, cumul, or normalize."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the calculation to perform on the data."))
	.Documentation()
);
LPOPERX WINAPI xll_quandl_calculate(xcstr trans)
{
#pragma XLLEXPORT
	static OPERX f, t(_T("&transformation="));

	try {
		f = _T("");

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
		}

		if (f != _T(""))
			f = XLL_XLF(Concatenate, t, f);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		f = OPERX(xlerr::NA);
	}

	return &f;
}

static AddInX xai_quandl_search(
	FunctionX(XLL_LPXLOPERX, _T("?xll_quandl_search"), _T("QUANDL.SEARCH"))
	.Arg(XLL_DOUBLEX, _T("Terms"), _T("one or more words to seach for."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Indicate the searching the data."))
	.Documentation()
);
LPXLOPERX WINAPI xll_quandl_search(LPOPERX term)
{
#pragma XLLEXPORT
	static OPERX query, space(_T(" ")), plus(_T("+"));
	
	try {
		query = _T("&query=");
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