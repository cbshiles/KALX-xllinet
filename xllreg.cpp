// xllreg.cpp - read and write JSON objects to the registry
#include "xlljson.h"
#include "xll/utility/registry.h"
#include <shellapi.h>

#pragma comment(lib, "shell32.lib")

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xstring xstring;

void xll_reg_set_object(const HKEY& h, xcstr k, xcstr sk, const OPERX& o);
void xll_reg_set_array(const HKEY& h, xcstr k, xcstr sk, const OPERX& o);

inline void xll_reg_set(const HKEY& h, xcstr k, const OPERX& o)
{
	Reg::CreateKey<xchar> rck(h, k);

	for (xword i = 0; i < o.rows(); ++i) {
		ensure (o(i,0).xltype == xltypeStr && o(i,0).val.str[0]);

		if (o(i,0).val.str[1] == '{') {
			xll_reg_set_object(h, k, o(i,0).string().c_str(), o(i,1));
		}
		else if (o(i,0).val.str[1] == '[') {
			xll_reg_set_array(h, k, o(i,0).string().c_str(), o(i,1));
		}
		else {
			xstring v = json::to_value(o(i,1));
			Reg::Object<xchar,xcstr> ro(h, k, o(i,0).string().c_str(), v.c_str());
		}
	}
}
inline void xll_reg_set_object(const HKEY& h, xcstr k, xcstr sk, const OPERX& o)
{
	xstring key = xstring(k);
	key.append(_T("\\"));
	key.append(sk);
	Reg::CreateKey<xchar> rck(h, key.c_str());

	xll_reg_set(h, key.c_str(), o.xltype == xltypeNum ? *h2p<OPERX>(o.val.num) : o);
}

// print with leading 0's
inline xstring itos(xword i, xword n)
{
	ensure (i < n);

	xstring s;
	xword d = 1 + log10(n == 1 ? 1 : n - 1);

	while (i) {
		s.insert(0, 1, i%10 + '0');
		i /= 10;
		--d;
	}

	while (d--)
		s.insert(0, 1, '0');

	return s;
}

inline void xll_reg_set_array(const HKEY& h, xcstr k, xcstr sk, const OPERX& o)
{
	xstring key = xstring(k);
	key.append(_T("\\"));
	key.append(sk);
	Reg::CreateKey<xchar> rck(h, key.c_str());
			
	for (xword j = 0; j < o.size(); ++j) {
		xstring keyj = key;
		keyj.append(_T("\\"));
		keyj.append(itos(j,o.size()));

		if (o[j].xltype == xltypeMulti) {
			XLL_INFO("JSON.REG.SET: array contains a non-scalar element");
		}
		else {
			RegSetValue(h, keyj.c_str(), REG_SZ, o[j].to_string().c_str(), 0);
		}
	}
}

static AddInX xai_json_reg_set(
	FunctionX(XLL_CSTRINGX, _T("?xll_json_reg_set"), _T("JSON.REG.SET"))
	.Arg(XLL_CSTRINGX, _T("Key"), _T("is the registry key to use as the base"))
	.Arg(XLL_HANDLEX, _T("Object"), _T("is a handle to a JSON.PARSE object to write"))
	.Category(_T("JSON"))
	.FunctionHelp(_T("Write Object to HKCU under Key and return Key."))
	.Documentation()
);
xcstr WINAPI xll_json_reg_set(xcstr key, HANDLEX h)
{
#pragma XLLEXPORT
	try {
		const OPERX& o = *h2p<OPERX>(h);

		xll_reg_set(HKEY_CURRENT_USER, key, o);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return key;
}

OPERX xll_reg_get_array(const HKEY&h, xcstr k);

inline OPERX xll_reg_get(const HKEY& h, xcstr k)
{
	OPERX o;
	OPERX kv(1, 2);

	Reg::OpenKey<xchar> rok(h, k);

	// name-value pairs
	xchar name[256], value[0x3FFF];
	DWORD nname = 256, nvalue = 0x3FFF, type;
	for (DWORD i = 0; ERROR_NO_MORE_ITEMS != RegEnumValue(rok, i, name, &nname, 0, &type, reinterpret_cast<LPBYTE>(value), &nvalue); ++i) {
		ensure (type == REG_SZ);
		kv[0] = name;
		xcstr v = value;
		kv[1] = json::parse_value<XLOPERX>(&v);
		o.push_back(kv);
		nname = 256;
		nvalue = 0x3FFF;
	}

	// sub keys
	DWORD nkey = 256;
	xchar key[256];
	for (DWORD i = 0; ERROR_NO_MORE_ITEMS != RegEnumKeyEx(rok, i, key, &nkey, 0, 0, 0, 0); ++i) {
		kv[0] = key;

		xstring k_(k);
		k_.append(_T("\\"));
		k_.append(key);

		if (*key == '[') {
			kv[1] = xll_reg_get_array(h, k_.c_str());
		}
		else {
			kv[1] = xll_reg_get(h, k_.c_str());
		}

		o.push_back(kv);
	}

	return o;
}
inline OPERX xll_reg_get_array(const HKEY&h, xcstr k)
{
	DWORD nkeys;

	Reg::OpenKey<xchar> rok(h, k);
	ensure (ERROR_SUCCESS == RegQueryInfoKey(rok, 0, 0, 0, &nkeys, 0, 0, 0, 0, 0, 0, 0));

	OPERX o(1, nkeys);

	xchar key[256];
	DWORD nkey = 256;
	for (DWORD j = 0; ERROR_NO_MORE_ITEMS != RegEnumKeyEx(rok, j, key, &nkey, 0, 0, 0, 0); ++j) {
		xstring keyj = k;
		keyj.append(_T("\\"));
		keyj.append(key);
		Reg::OpenKey<xchar> rokj(h, keyj.c_str());

		DWORD i = _ttoi(key);
		o[i] = xll_reg_get(h, keyj.c_str());
		nkey = 256;
	}

	return o;
}


static AddInX xai_json_reg_get(
	FunctionX(XLL_HANDLEX, _T("?xll_json_reg_get"), _T("JSON.REG.GET"))
	.Arg(XLL_CSTRINGX, _T("Key"), _T("is the registry key to use as the base"))
	.Uncalced()
	.Category(_T("JSON"))
	.FunctionHelp(_T("Read JSON object from HKCU under Key."))
	.Documentation()
);
HANDLEX WINAPI xll_json_reg_get(xcstr key)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<OPERX> ho = new OPERX(xll_reg_get(HKEY_CURRENT_USER, key));

		h = ho.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

int
xll_regedit_close(void)
{
	try {
		if (Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Adjust")))
			Excel<XLOPER>(xlfDeleteCommand, OPER(7), OPER(4), OPER("Adjust"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		
		return 0;
	}

	return 1;
}
static Auto<Close> xac_regedit(xll_regedit_close);

static AddIn xai_regedit("?xll_regedit", "XLL.REGEDIT");
int WINAPI xll_regedit(void)
{
#pragma XLLEXPORT
	try {
		OPERX ac = XLL_XLF(ActiveCell);
		OPERX co = XLL_XL_(Coerce, ac);
		ensure (co.xltype == xltypeStr);

		Reg::Object<xchar, xcstr> lk(HKEY_CURRENT_USER, _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit"), _T("LastKey"));
		lk = co.string().c_str();
		HINSTANCE h = ShellExecute(0, _T("open"), _T("regedit"), 0, 0, SW_SHOW);
		ensure ((int)h > 32);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

int
xll_regedit_open(void)
{
	try {
		// Try to add just above first menu separator.
		OPER oPos;
//		oPos = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("-"));
		oPos = 7;

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
static Auto<Open> xao_regedit_open(xll_regedit_open);


#ifdef _DEBUG

void xll_test_reg_itos()
{
	ensure (itos(0, 1) == _T("0"));
	ensure (itos(9, 10) == _T("9"));
	ensure (itos(0, 11) == _T("00"));
	ensure (itos(12, 100) == _T("12"));
	ensure (itos(12, 101) == _T("012"));
}

int xll_test_reg()
{
	try {
		xll_test_reg_itos();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
static Auto<OpenAfterX> xao_test_reg(xll_test_reg);

#endif // _DEBUG