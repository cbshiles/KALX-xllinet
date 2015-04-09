// xllfwf.cpp - Fixed Width Format parsing
#include "xllinet.h"

using namespace xll;

static AddInX xai_fwf_get(
	FunctionX(XLL_LPOPERX, _T("?xll_fwf_get"), _T("FWF.GET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a string"))
	.Arg(XLL_FPX, _T("Widths"), _T("is an array of widths to use. Negative widths are discarded."))
	.Arg(XLL_BOOLX, _T("_Raw"), _T("is an optional boolean value indicating raw strings are to be returned."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Parse a string associated with Handle into fixed width fields."))
	.Documentation()
);
LPOPERX xll_fwf_get(HANDLEX h, xfp* pw, BOOL raw)
{
#pragma XLLEXPORT
	using st = xstring::size_type;
	static OPERX o;

	try {
		o.resize(0,0);
		handle<xstring> hs(h);

		st pos, prev = 0;
		while ((pos = hs->find('\n', prev)) != xstring::npos) {
			OPERX row;			
			st off = prev;
			for (xword i = 0; i < size(*pw); ++i) {
				xstring::size_type wi = static_cast<st>(fabs(pw->array[i]));
				if (pw->array[i] >= 0) {
					if (off + wi <= pos)
						row.push_back(hs->substr(off, wi).c_str());
					else
						row.push_back(_T(""));
				}
				off += wi;
			}
			o.push_back(row.resize(1, row.size()));
			
			prev = pos + 1;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	return &o;
}