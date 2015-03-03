// xllfile.cpp - read and write files
#define EXCEL12
#include "xll/xll.h"
#include "xll/utility/scoped_handle.h"
#include "xllcsv.h"
#include "xlljson.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xstring xstring;

static AddInX xai_file_read(
	FunctionX(XLL_HANDLEX, _T("?xll_file_read"), _T("FILE.READ"))
	.Arg(XLL_CSTRINGX, _T("File"), _T("is the path of the file to read."))
	.Uncalced()
	.Category(_T("XLL"))
	.FunctionHelp(_T("Return a handle to a string containing the file."))
	.Documentation()
);
HANDLEX WINAPI xll_file_read(xcstr file)
{
#pragma XLLEXPORT
	handlex h;

	try {
		scoped_handle::base<HANDLE> f(CreateFile(file, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0));

		DWORD hi, lo = GetFileSize(f,&hi);
		ensure (hi == 0);

		handle<xstring> hs(new xstring());
		hs->resize(lo);
		ensure (ReadFile(f, &hs->operator[](0), lo, &hi, 0));
		ensure (lo == hi);
	} 
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddInX xai_file_write(
	FunctionX(XLL_DOUBLEX, _T("?xll_file_write"), _T("FILE.WRITE"))
	.Arg(XLL_CSTRINGX, _T("File"), _T("is the path of the file to write."))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is the range to be written to File."))
	.Arg(XLL_CSTRINGX, _T("_Method"), _T("is either \"CSV\" or \"JSON\". Default is CSV."))
	.Uncalced()
	.Category(_T("XLL"))
	.FunctionHelp(_T("Return then number of bytes written."))
	.Documentation()
);
double WINAPI xll_file_write(xcstr file, LPOPERX po, xcstr method)
{
#pragma XLLEXPORT
	DWORD n(0);

	try {
		scoped_handle::base<HANDLE> f(CreateFile(file, GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0));

		xstring s;
		if (::tolower(method[0]) == 'j')
			s = json::to_object(*po);
		else
			s = csv::to_string(*po);

		WriteFile(f, &s[0], s.length()*sizeof(xchar), &n, 0);
	} 
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return n;
}

#ifdef _DEBUG

int xll_test_file()
{
	try {
		OPERX file = _T("test.csv");
		OPERX o  {
			OPERX(1.23), OPERX(_T("foo")),
			OPERX(), OPERX(xlerr::NA)
		};
		o.resize(2,2);

		Excel<XLOPERX>(xlUDF, OPERX(_T("FILE.WRITE")), file, o);
		OPERX o_ = Excel<XLOPERX>(xlUDF, OPERX(_T("FILE.READ")), file);
//		ensure (o_ == o);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

//static Auto<OpenAfterX> xao_test_file(xll_test_file);

#endif // _DEBUG