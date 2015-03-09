// xllcsv.h - comma separated values
#pragma once
//#define EXCEL12
#include "xll/xll.h"

#define CATEGORY _T("XLL")

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;

namespace csv {

	inline void trim(xcstr* pb, xcstr* pe = 0)
	{
		ensure(pb && *pb);
		while (::_istspace(**pb) && *pb < *pe)
			++*pb;
		if (pe && *pe)
			while (::_istspace((*pe)[-1]) && *pb < *pe)
				--*pe;
	}
	inline void unquote(xcstr *pb, xcstr *pe) {
		if (**pb == '"' && (*pe)[-1] == '"') {
			++*pb;
			--*pe;
		}
	}
	// scan for next unquoted fs
	inline xcstr scan(xcstr s, xcstr t, xcstr fs, xcstr q = _T("\"\'"))
	{
		xcstr inq = 0;

		for (; s && *s && s < t; ++s) {
			xcstr inq_ = _tcschr(q, *s);
			
			if (!inq)
				inq = inq_;
			else if (inq == inq_)
				inq = 0; // closing quote

			if (!inq && _tcschr(fs, *s)) {
				break;
			}
		}

		return s;
	}
	inline OPERX parse(xcstr s, xcstr fs = _T(","), xcstr rs = _T(";\n"))
	{
		OPERX o;

		for (xcstr t = s, u = _tcspbrk(t, rs); u; u = _tcspbrk(t = u + 1, rs)) {
			OPERX row;
			// parse record [t, u)
			for (xcstr v = t, w = scan(v, u, fs); v <= u; w = scan(v = w + 1, u, fs)) {
				row.push_back(to_XOPER<XLOPERX>(v, w - v));
			}

			o.push_back(row.resize(1, row.size()));
		}

		return o;
	}

	inline xstring to_string(const OPERX& o, xcstr fs = _T(","), xcstr rs = _T("\n"))
	{
		xstring s;

		for (xword i = 0; i < o.rows(); ++i) {
			s.append(o(i, 0).to_string());
			for (xword j = 1; j < o.columns(); ++j) {
				s.append(fs);
				s.append(o(i,j).to_string());
			}
			s.append(rs);
		}

		return s;
	}
}
