// json.h - JSON objects for Excel
/*
k1 | v1
k2 | {k3 | v3         -- nested objects
      k4 | v4}
k5 | [v6 | v7 | v8 ]  -- array is one row (what if 2 elements???)
*/
#pragma once
#define EXCEL12
#include "xll/xll.h"

#define CATEGORY_JSON _T("JSON")

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;

namespace json {

	template<class X>
	inline typename xll::traits<X>::xcstr eat(typename xll::traits<X>::xcstr str, int c)
	{
		str = skip_white<X>(str);
		
		ensure (*str++ == c);
		
		return str;
	}
	// non throwing eat
	template<class X>
	inline bool peck(typename xll::traits<X>::xcstr* pstr, int c)
	{
		*pstr = skip_white<X>(*pstr);

		return **pstr == c ? ++*pstr, true : false;
	}

	template<class X>
	inline bool is_white(typename xll::traits<X>::xchar c)
	{
		return c == ' ' || c == '\t' || c == '\n' || c == '\r';
	}
	template<class X>
	inline typename xll::traits<X>::xcstr skip_white(typename xll::traits<X>::xcstr str)
	{
		while (is_white<X>(*str))
			++str;

		return str;
	}
	template<class X>
	inline bool is_equal(typename xll::traits<X>::xcstr c, int* s)
	{
		while (*c && *s && *c++ == *s++)
			;

		return *s == 0;
	}
	template<class X>
	inline bool is_null(typename xll::traits<X>::xcstr c)
	{
		static int null[] = {'n','u','l','l',0};

		return is_equal<X>(c, null);
	}
	template<class X>
	inline bool is_true(typename xll::traits<X>::xcstr c)
	{
		static int true_[] = {'t','r','u','e',0};

		return is_equal<X>(c, true_);
	}
	template<class X>
	inline bool is_false(typename xll::traits<X>::xcstr c)
	{
		static int false_[] = {'f','a','l','s','e',0};

		return is_equal<X>(c, false_);
	}
	template<class X>
	inline typename xll::traits<X>::xchar len(typename xll::traits<X>::xcstr b, typename xll::traits<X>::xcstr e)
	{
		return static_cast<xll::traits<X>::xchar>(e - b);
	}

	template<class X>
	inline XOPER<X> parse_number(typename xll::traits<X>::xcstr* pstr)
	{
		ensure (isdigit(**pstr) || **pstr == '+' || **pstr == '-' || **pstr == '.');

		double d = xll::traits<X>::strtod(*pstr, pstr);

		return XOPER<X>(d);
	}

	// string: " chars "
	template<class X>
	inline XOPER<X> parse_string(typename xll::traits<X>::xcstr* pstr)
	{
		OPERX o;
		xll::traits<X>::xcstr b, e;

		ensure (**pstr == '\"');
		b = *pstr + 1;
		for (e = b; *e != 0 && *e != '\"'; ++e)
			if (*e == '\\' && *(e + 1) == '\"')
				++e;

		ensure (*e == '\"');

		*pstr = e + 1;

		o = XOPER<X>(b, len<X>(b, e));

		return o;
	}

	// elements: value or value , elements
	template<class X>
	inline XOPER<X> parse_elements(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o;

		o = parse_value<X>(pstr);
		while (peck<X>(pstr, ','))
			o.push_back(parse_value<X>(pstr));

		return o;
	}
	// array: [ ], [ elements ]
	template<class X>
	inline XOPER<X> parse_array(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o;

		*pstr = eat<X>(*pstr, '[');

		if (peck<X>(pstr, ']'))
			return o; // empty array
		
		o = parse_elements<X>(pstr);
		o.resize(1, o.size());

		*pstr = eat<X>(*pstr, ']');

		return o;
	}
	// value: string, number, object, array, true, false, null
	template<class X>
	inline XOPER<X> parse_value(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o;
		xcstr b, e;

		*pstr = skip_white<X>(*pstr);

		b = e = *pstr;
		if (*b == '{') {
			o = handle<XOPER<X>>(new XOPER<X>(parse_object<X>(&e))).get();
		}
		else if (*b == '[') {
			o = handle<XOPER<X>>(new XOPER<X>(parse_array<X>(&e))).get();
		}
		else if (*b == '"') {
			o = parse_string<X>(&e);
		}
		else if (is_null<X>(b)) {
			e = b + 4;
			o = XOPER<X>(xlerr::Null);
		}
		else if (is_true<X>(b)) {
			e = b + 4;
			o = XOPER<X>(true);
		}
		else if (is_false<X>(b)) {
			e = b + 5;
			o = XOPER<X>(false);
		}
		else {
			o = parse_number<X>(&e);
		}
		
		*pstr = e;

		return o;
	}
	// pair: string ':' value
	template<class X>
	inline XOPER<X> parse_pair(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o(1, 2);

		*pstr = skip_white<X>(*pstr);
		o[0] = parse_string<X>(pstr);
		*pstr = eat<X>(*pstr, ':');
		o[1] = parse_value<X>(pstr);

		// handle keys start with '*'
		if (o[1].xltype == xltypeNum && xll::handle<XOPER<X>>(o[1].val.num, false)) {
			xll::traits<X>::xchar c = '*';
			o[0] = XLL_XLF(Concatenate, XOPER<X>(&c, 1), o[0]);
		}

		return o;
	}
	// members: pair or pair , members
	template<class X>
	inline XOPER<X> parse_members(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o;

		o.push_back(parse_pair<X>(pstr));
		while (peck<X>(pstr, ','))
			o.push_back(parse_pair<X>(pstr));

		return o;
	}
	// object: { }, { members }
	template<class X>
	inline XOPER<X> parse_object(typename xll::traits<X>::xcstr* pstr)
	{
		XOPER<X> o;

		*pstr = eat<X>(*pstr, '{');

		if (peck<X>(pstr, '}'))
			return o; // empty object

		o = parse_members<X>(pstr);
		*pstr = eat<X>(*pstr, '}');

		return o;
	}

	// stringify
	template<class X>
	inline typename xll::traits<X>::xstring to_elements(const XOPER<X>& e)
	{
		xll::traits<X>::xstring str;

		str.append(to_value(e[0]));
		for (xword i = 1; i < e.size(); ++i) {
			str.append(1, ',').append(to_value(e[i]));
		}

		return str;
	}

	template<class X>
	inline typename xll::traits<X>::xstring to_value(const XOPER<X>& v)
	{
		xll::traits<X>::xstring str;

		if (v.xltype == xltypeErr || v.xltype == xltypeMissing || v.xltype == xltypeNil) {
			str = _T("null");
		}
		else if (v.size() == 1 && (v.xltype != xltypeMulti)) {
			if (v.xltype == xltypeStr) {
				str.append(1, '\"').append(v.val.str + 1, v.val.str[0]).append(1, '\"');
			}
			else if (v.xltype == xltypeBool) {
				str = xll::traits<X>::string(v.val.xbool ? "true" : "false");
			}
			else if (v.xltype == xltypeNum) {
				xll::handle<XOPER<X>> h(v.val.num, false);
				if (h) {
					str = to_value(*h);
				}
				else {
					str = xll::traits<X>::to_string(v.val.num);
				}
			}
			else if (v.rows() == 1) {
				str.append(1, '[').append(to_elements(v)).append(1, ']');
			}
			else {
				str.append(1, '{').append(to_object(v)).append(1, '}');
			}
		}
		else { // object or array
			ensure (v.size() > 1);

			if (v.rows() == 1) {
				str.append(1, '[');
				str.append(to_elements(v));
				str.append(1, ']');
			}
			else {
				str.append(to_object(v));
			}
		}

		return str;
	}

		
	template<class X>
	inline typename xll::traits<X>::xstring to_pair(const XOPER<X>& k, const XOPER<X>& v)
	{
		ensure (k.xltype == xltypeStr);

		return xstring(1, '"').append(k.val.str + 1, k.val.str[0]).append(1, '"').append(1, ':').append(to_value(v));
	}

	template<class X>
	inline typename xll::traits<X>::xstring to_member(const XOPER<X>& o)
	{
		xll::traits<X>::xstring str;

		if (o.size() == 0)
			return str;

		ensure (o.columns() == 2);

		str.append(to_pair(o[0], o[1]));
		for (xword i = 1; i < o.rows(); ++i) {
			str.append(1, ',').append(to_pair(o(i, 0), o(i, 1)));
		}

		return str;
	}
	template<class X>
	inline typename xll::traits<X>::xstring to_object(const XOPER<X>& o)
	{
		return xll::traits<X>::xstring(1, '{').append(to_member(o)).append(1, '}');
	}
	// JSON OPERX to string
	template<class X>
	inline typename xll::traits<X>::xstring stringify(const XOPER<X>& o)
	{
		return to_object(o);
	}

} // namespace JSON