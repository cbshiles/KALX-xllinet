// xlljson.cpp - JSON objects in Excel
/*
JSON.PARSE converts a string into a 2 column recursive object.
JSON.GET converts a handle to the object into a two column array
Arrays and objects get replace by the lparray pointer. These
are placed in a set as they are displayed.
JSON.GET also converts lparray pointers to the appropriate ranges
*/
#include <set>
#include "xlljson.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xstring xstring;
typedef traits<XLOPERX>::xword xword;

static std::set<HANDLEX> json_subobject;
// return JSON object or subobject
const XLOPERX& json_handle(HANDLEX h)
{
	LPXLOPERX ph = h2p<XLOPERX>(h);
	if (json_subobject.find(h) != json_subobject.end()) {
		return *ph;
	}

	return *handle<OPERX>(h);	
}

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

static AddInX xai_json_parse(
	FunctionX(XLL_HANDLEX, _T("?xll_json_parse"), _T("JSON.PARSE"))
	.Arg(XLL_LPOPERX, _T("Handle"), _T("is a handle to a JSON object"))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a handle to a JSON object"))
	.Documentation()
	);
HANDLEX WINAPI xll_json_parse(LPOPERX po)
{
#pragma XLLEXPORT
	handlex h;

	try {
		OPERX o;

		if (po->xltype == xltypeNum) {
			handle<xstring> hs(po->val.num);
			xcstr cs = hs->c_str();
			o = json::parse_object<XLOPERX>(&cs);
		}
		else if (po->xltype == xltypeStr) {
			xcstr cs = po->val.str + 1;
			o = json::parse_object<XLOPERX>(&cs); //!!! malformed objects???
		}
		else {
			XLL_WARNING("JSON.PARSE: argument must be a string or a handle to a string");
			h = 0;
		}

		if (h != 0)
			h = handle<OPERX>(new OPERX(o)).get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		h = 0;
	}
	catch (...) {
		XLL_ERROR("JSON.GET: unknown exception");

		h = 0;
	}

	return h;
}

static AddInX xai_json_get(
	FunctionX(XLL_LPOPERX, _T("?xll_json_get"), _T("JSON.GET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a parsed JSON object"))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Return a JSON object"))
	.Documentation()
);
LPOPERX WINAPI xll_json_get(HANDLEX h)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		LPXLOPERX ph = h2p<XLOPERX>(h);

		o = json_handle(h);

		// replace multi's by handles
		for (xword i = 0; i < o.size(); ++i) {
			if (o[i].xltype == xltypeMulti) {
				HANDLEX hx = p2h<XLOPERX>(ph->val.array.lparray + i);
				o[i] = hx;
				json_subobject.insert(hx);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}
	catch (...) {
		XLL_ERROR("JSON.GET: unknown exception");

		o = OPERX(xlerr::NA);
	}

	return &o;
}

// match a with initial substring b, ignoring asterisks
inline int json_match(const OPERX& a, const OPERX& b)
{
	ensure (a.xltype == xltypeStr);
	ensure (b.xltype == xltypeStr);

	xchar na = a.val.str[0];
	xchar nb = b.val.str[0];

	xcstr pa = a.val.str;
	xcstr pb = b.val.str;

	if (na > 0 && pa[1] == '*') {
		++pa;
		--na;
	}
	if (nb > 0 && pb[1] == '*') {
		++pb;
		--nb;
	}

	while (na && nb && (::tolower(*++pa) == ::tolower(*++pb))) {
		--na;
		--nb;
	}

	return nb == 0 ? na : -nb;
}

static AddInX xai_json_value(
	FunctionX(XLL_LPOPERX, _T("?xll_json_value"), _T("JSON.VALUE"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a range."))
	.Arg(XLL_LPOPERX, _T("Key"), _T("is the object key to lookup."))
	.Category(CATEGORY_JSON)
	.FunctionHelp(_T("Return the value corresponding to Key."))
	.Documentation(_T("If the value is an array or object and Key starts with asterisk, then the handle is returned."))
);
LPOPERX WINAPI xll_json_value(HANDLEX h, LPOPERX pk)
{
#pragma XLLEXPORT
	static OPERX v;

	try {
		v = OPERX(xlerr::NA);

		const OPERX& o = json_handle(h);
		ensure (o.columns() == 2);
		
		const OPERX& k(*pk);
		ensure (k.size() == 1);

		int n_ = 0;
		xword i_ = (xword)-1;
		for (xword i = 0; i < o.rows(); ++i) {
			int n = json_match(o(i, 0), k);
			if (n == 0) {
				i_ = i;

				break;
			}
			if (n_ < n) {
				i_ = i; // best match
				n_ = n;
			}
		}

		if (i_ != (xword)-1)
			v = o(i_,1); // don't copy!!! return address

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		v = OPERX(xlerr::NA);
	}

	return &v;
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
		OPERX off;
		if (ho->xltype == xltypeNil) {
			off = XLL_XLF(Offset, ac, OPERX(0), OPERX(1));
			XLL_XL_(Set, off, OPERX(xlerr::Null));
		}
		else {
			off = XLL_XLF(Offset, ac, OPERX(0), OPERX(1), OPERX(ho->rows()), OPERX(ho->columns()));
			XLL_XL_(Set, off, *ho);
		}
		XLL_XLC(Select, off);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("XLL.EXPAND: unknown exception");

		return FALSE;
	}

	return TRUE;
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

#ifdef _DEBUG

void xll_test_json_match()
{
	int n;

	n = json_match(OPERX(_T("abc")), OPERX(_T("abc")));
	ensure (n == 0);

	n = json_match(OPERX(_T("*abc")), OPERX(_T("abc")));
	ensure (n == 0);

	n = json_match(OPERX(_T("abc")), OPERX(_T("*abc")));
	ensure (n == 0);

	n = json_match(OPERX(_T("*abc")), OPERX(_T("*abc")));
	ensure (n == 0);

	n = json_match(OPERX(_T("*ab")), OPERX(_T("*abc")));
	ensure(n == -1);

	n = json_match(OPERX(_T("ab")), OPERX(_T("*abc")));
	ensure(n == -1);

	n = json_match(OPERX(_T("*ab")), OPERX(_T("abc")));
	ensure(n == -1);

	n = json_match(OPERX(_T("ab")), OPERX(_T("abc")));
	ensure(n == -1);
}
void xll_test_json_value()
{
	OPERX o{
		OPERX(_T("abc")), OPERX(1),
		OPERX(_T("ab")), OPERX(2),
		OPERX(_T("a")), OPERX(3)
	};
	o.resize(3, 2);
	handle<OPERX> ho = new OPERX(o);

	OPERX k, v;
	k = _T("a");
	v = *xll_json_value(ho.get(), &k);
	ensure (v == 3);

	k = _T("ab");
	v = *xll_json_value(ho.get(), &k);
	ensure (v == 2);

	k = _T("abc");
	v = *xll_json_value(ho.get(), &k);
	ensure (v == 1);

	k = _T("xyz");
	v = *xll_json_value(ho.get(), &k);
	ensure (v == OPERX(xlerr::NA));
}
void xll_test_json_parse(void)
{
//	xcstr text = _T("{\"help\": \"Return...\", \"success\" : true, \"result\" : {\"license_title\": \"Creative Commons Attribution\", \"maintainer\" : \"\", \"relationships_as_object\" : [], \"private\" : false}}");
	xcstr text = _T("{\"maintainer\" : \"\"}");
	OPERX o;
	o = json::parse_object<XLOPERX>(&text);
}

int xll_test_json()
{
	try {
		test_json();
		xll_test_json_match();
		xll_test_json_value();
		xll_test_json_parse();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
static Auto<OpenAfterX> xao_test_json(xll_test_json);

#endif // _DEBUG