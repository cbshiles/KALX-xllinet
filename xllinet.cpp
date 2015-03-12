// xllinet.cpp - WinInet wrappers
#include "xllinet.h"
#include "xll/utility/inet.h"

#define CATEGORY_ENUM "Inet Enumerations"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xstring xstring;

static Inet::Open io; // default

static AddInX xai_inet_open(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_open"), _T("INET.OPEN"))
	.Arg(XLL_CSTRINGX, _T("_Agent"), _T("is the option user agent. Default is \"Mozilla/5.0\"."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to an Inet::Open object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_inet_open(xcstr agent)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<Inet::Open> hio = new Inet::Open(agent ? agent : _T("Mozilla/5.0"));

		h = hio.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

XLL_ENUM(INTERNET_DEFAULT_FTP_PORT,INTERNET_DEFAULT_FTP_PORT, CATEGORY_ENUM, "Uses the default port for FTP servers (port 21).");
XLL_ENUM(INTERNET_DEFAULT_GOPHER_PORT,INTERNET_DEFAULT_GOPHER_PORT, CATEGORY_ENUM, "Uses the default port for Gopher servers (port 70).");
XLL_ENUM(INTERNET_DEFAULT_HTTP_PORT,INTERNET_DEFAULT_HTTP_PORT, CATEGORY_ENUM, "Uses the default port for HTTP servers (port 80).");
XLL_ENUM(INTERNET_DEFAULT_HTTPS_PORT,INTERNET_DEFAULT_HTTPS_PORT, CATEGORY_ENUM, "Uses the default port for Secure Hypertext Transfer Protocol (HTTPS) servers (port 443).");
XLL_ENUM(INTERNET_DEFAULT_SOCKS_PORT,INTERNET_DEFAULT_SOCKS_PORT, CATEGORY_ENUM, "Uses the default port for SOCKS firewall servers (port 1080).");
XLL_ENUM(INTERNET_INVALID_PORT_NUMBER,INTERNET_INVALID_PORT_NUMBER, CATEGORY_ENUM, "Uses the default port for the service specified by dwService.");

XLL_ENUM(INTERNET_SERVICE_FTP,INTERNET_SERVICE_FTP, CATEGORY_ENUM, "FTP service.");
XLL_ENUM(INTERNET_SERVICE_GOPHER,INTERNET_SERVICE_GOPHER, CATEGORY_ENUM, "Gopher service.");
XLL_ENUM(INTERNET_SERVICE_HTTP,INTERNET_SERVICE_HTTP, CATEGORY_ENUM, "HTTP service.");

static AddInX xai_inet_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_connect"), _T("INET.CONNECTION"))
	.Arg(XLL_HANDLEX, _T("Open"), _T("is an optional handle returned by INET.OPEN"))
	.Arg(XLL_CSTRINGX, _T("Server"), _T("is the name or IP address of the server"))
	.Arg(XLL_SHORTX, _T("Port"), _T("is the optional port on the server. See INTERNET_DEFAULT_*"))
	.Arg(XLL_CSTRINGX, _T("User"), _T("is the optional user name"))
	.Arg(XLL_CSTRINGX, _T("Pass"), _T("is the optional password."))
	.Arg(XLL_SHORTX, _T("Service"), _T("is the type of service to access. See INTERNET_SERVICE_*."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to an Inet::Open::Connect object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_inet_connect(HANDLEX io, xcstr server, SHORT port, xcstr user, xcstr pass, SHORT service)
{
#pragma XLLEXPORT
	handlex h;

	try {
		if (!port)
			port = INTERNET_DEFAULT_HTTP_PORT;
		if (!service)
			service = INTERNET_SERVICE_HTTP;

		handle<Inet::Open> hio(io);
		handle<Inet::Open::Connection> hioc = new Inet::Open::Connection(*hio, server, port, user, pass, service);

		h = hioc.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddInX xai_inet_request(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_request"), _T("INET.REQUEST"))
	.Arg(XLL_HANDLEX, _T("Connect"), _T("is a handle returned by INET.CONNECTION"))
	.Arg(XLL_CSTRINGX, _T("Verb"), _T("is the optional HTTP verb to use in the request. Default is GET."))
	.Arg(XLL_CSTRINGX, _T("Object"), _T("is the optional object name of the HTTP verb."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to an Inet::Open::Connection::Request object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_inet_request(HANDLEX ioc, xcstr verb, xcstr object)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<Inet::Open::Connection> hioc(ioc);
		handle<Inet::Open::Connection::Request> hiocr = new Inet::Open::Connection::Request(*hioc, verb, object);

		h = hiocr.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

XLL_ENUM(HTTP_ADDREQ_FLAG_ADD, HTTP_ADDREQ_FLAG_ADD, CATEGORY_ENUM, "Adds the header if it does not exist. Used with HTTP_ADDREQ_FLAG_REPLACE.");
XLL_ENUM(HTTP_ADDREQ_FLAG_ADD_IF_NEW, HTTP_ADDREQ_FLAG_ADD_IF_NEW, CATEGORY_ENUM, "Adds the header only if it does not already exist; otherwise, an error is returned.");
XLL_ENUM(HTTP_ADDREQ_FLAG_COALESCE, HTTP_ADDREQ_FLAG_COALESCE, CATEGORY_ENUM, "Coalesces headers of the same name.");
XLL_ENUM(HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA, HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA, CATEGORY_ENUM, "Coalesces headers of the same name.");
XLL_ENUM(HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON, HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON, CATEGORY_ENUM, "Coalesces headers of the same name using a semicolon.");
XLL_ENUM(HTTP_ADDREQ_FLAG_REPLACE, HTTP_ADDREQ_FLAG_REPLACE, CATEGORY_ENUM, "Replaces or removes a header. If the header value is empty and the header is found, it is removed. If not empty, the header value is replaced.");

static AddInX xai_inet_add_request_headers(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_add_request_headers"), _T("INET.REQUEST.HEADERS"))
	.Arg(XLL_HANDLEX, _T("Request"), _T("is a handle returned by INET.REQUEST"))
	.Arg(XLL_CSTRINGX, _T("Header"), _T("is the header to add to the request."))
	.Arg(XLL_WORDX, _T("Flags"), _T("are optional flags from HTTP_ADDREQ_FLAG_*."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to the Inet::Open::Connection::Request object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_inet_add_request_headers(HANDLEX iocr, xcstr headers, WORD flags)
{
#pragma XLLEXPORT
	try {
		handle<Inet::Open::Connection::Request>(iocr)->AddHeader(headers);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return iocr;
}

static AddInX xai_inet_request_send(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_request_send"), _T("INET.REQUEST.SEND"))
	.Arg(XLL_HANDLEX, _T("Request"), _T("is a handle returned by INET.REQUEST"))
	.Arg(XLL_CSTRINGX, _T("Header"), _T("is the header to add to the request."))
	.Arg(XLL_CSTRINGX, _T("Optional"), _T("is a string of optional data to send"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to the Inet::Open::Connection::Request object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_inet_request_send(HANDLEX iocr, xcstr headers, xcstr optional)
{
#pragma XLLEXPORT
	try {
		handle<Inet::Open::Connection::Request> hiocr(iocr);
		DWORD hlen = static_cast<DWORD>(_tcslen(headers));
		DWORD olen = static_cast<DWORD>(_tcslen(optional));

		HttpSendRequest(*hiocr, headers, hlen, (LPVOID)optional, olen);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return iocr;
}

XLL_ENUM(INTERNET_FLAG_EXISTING_CONNECT, INTERNET_FLAG_EXISTING_CONNECT, CATEGORY_ENUM, "Attempts to use an existing InternetConnect object if one exists with the same attributes required to make the request.");
XLL_ENUM(INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_HYPERLINK, CATEGORY_ENUM, "Forces a reload if there was no Expires time and no LastModified time returned from the server when determining whether to reload the item from the network.");
XLL_ENUM(INTERNET_FLAG_IGNORE_CERT_CN_INVALID, INTERNET_FLAG_IGNORE_CERT_CN_INVALID, CATEGORY_ENUM, "Disables checking of SSL/PCT-based certificates that are returned from the server against the host name given in the request.");
XLL_ENUM(INTERNET_FLAG_IGNORE_CERT_DATE_INVALID, INTERNET_FLAG_IGNORE_CERT_DATE_INVALID, CATEGORY_ENUM, "Disables checking of SSL/PCT-based certificates for proper validity dates.");
XLL_ENUM(INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP, INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP, CATEGORY_ENUM, "Disables detection of this special type of redirect. When this flag is used, WinINet transparently allows redirects from HTTPS to HTTP URLs.");
XLL_ENUM(INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS, INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS, CATEGORY_ENUM, "Disables the detection of this special type of redirect. When this flag is used, WinINet transparently allows redirects from HTTP to HTTPS URLs.");
XLL_ENUM(INTERNET_FLAG_KEEP_CONNECTION, INTERNET_FLAG_KEEP_CONNECTION, CATEGORY_ENUM, "Uses keep-alive semantics, if available, for the connection. This flag is required for Microsoft Network (MSN), NTLM, and other types of authentication.");
XLL_ENUM(INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_NEED_FILE, CATEGORY_ENUM, "Causes a temporary file to be created if the file cannot be cached.");
XLL_ENUM(INTERNET_FLAG_NO_AUTH, INTERNET_FLAG_NO_AUTH, CATEGORY_ENUM, "Does not attempt authentication automatically.");
XLL_ENUM(INTERNET_FLAG_NO_AUTO_REDIRECT, INTERNET_FLAG_NO_AUTO_REDIRECT, CATEGORY_ENUM, "Does not automatically handle redirection in HttpSendRequest.");
XLL_ENUM(INTERNET_FLAG_NO_CACHE_WRITE, INTERNET_FLAG_NO_CACHE_WRITE, CATEGORY_ENUM, "Does not add the returned entity to the cache.");
XLL_ENUM(INTERNET_FLAG_NO_COOKIES, INTERNET_FLAG_NO_COOKIES, CATEGORY_ENUM, "Does not automatically add cookie headers to requests, and does not automatically add returned cookies to the cookie database.");
XLL_ENUM(INTERNET_FLAG_NO_UI, INTERNET_FLAG_NO_UI, CATEGORY_ENUM, "Disables the cookie dialog box.");
XLL_ENUM(INTERNET_FLAG_PASSIVE, INTERNET_FLAG_PASSIVE, CATEGORY_ENUM, "Uses passive FTP semantics. InternetOpenUrl uses this flag for FTP files and directories.");
XLL_ENUM(INTERNET_FLAG_PRAGMA_NOCACHE, INTERNET_FLAG_PRAGMA_NOCACHE, CATEGORY_ENUM, "Forces the request to be resolved by the origin server, even if a cached copy exists on the proxy.");
XLL_ENUM(INTERNET_FLAG_RAW_DATA, INTERNET_FLAG_RAW_DATA, CATEGORY_ENUM, "Returns the data as a WIN32_FIND_DATA structure when retrieving FTP directory information.");
XLL_ENUM(INTERNET_FLAG_SECURE, INTERNET_FLAG_SECURE, CATEGORY_ENUM, "Uses secure transaction semantics. This translates to using Secure Sockets Layer/Private Communications Technology (SSL/PCT) and is only meaningful in HTTP requests.");

static AddInX xai_inet_url(
	FunctionX(XLL_HANDLEX, _T("?xll_inet_url"), _T("INET.URL"))
	.Arg(XLL_CSTRINGX, _T("Url"), _T("is the URL to connect with."))
	.Arg(XLL_CSTRINGX, _T("Headers"), _T("are optional headers to send with the request."))
	.Arg(XLL_WORDX, _T("Flags"), _T("are optional flags from INTERNET_FLAG_*."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to data at the Url."))
	.Documentation(_T("Functional replacement for WEBSERVICE except it returns a string handle."))
);
HANDLEX WINAPI xll_inet_url(xcstr url, xcstr headers, WORD flags)
{
#pragma XLLEXPORT
	handlex h;

	try {
		DWORD len = static_cast<DWORD>(_tcslen(headers));
		Inet::Open::URL iou(io, url, headers, len, flags);
		handle<xstring> hs = new xstring(traits<XLOPERX>::string(iou.Read().c_str()));

		h = hs.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

XLL_ENUM(ICU_BROWSER_MODE,ICU_BROWSER_MODE,CATEGORY_ENUM,"Does not encode or decode characters after \"#\" or \"?\", and does not remove trailing white space after \"?\". If this value is not specified, the entire URL is encoded and trailing white space is removed.");
XLL_ENUM(ICU_DECODE,ICU_DECODE,CATEGORY_ENUM,"Converts all %XX sequences to characters, including escape sequences, before the URL is parsed.");
XLL_ENUM(ICU_ENCODE_PERCENT,ICU_ENCODE_PERCENT,CATEGORY_ENUM,"Encodes any percent signs encountered. By default, percent signs are not encoded. This value is available in Microsoft Internet Explorer 5 and later.");
XLL_ENUM(ICU_ENCODE_SPACES_ONLY,ICU_ENCODE_SPACES_ONLY,CATEGORY_ENUM,"Encodes spaces only.");
XLL_ENUM(ICU_NO_ENCODE,ICU_NO_ENCODE,CATEGORY_ENUM,"Does not convert unsafe characters to escape sequences.");
XLL_ENUM(ICU_NO_META,ICU_NO_META,CATEGORY_ENUM,"Does not remove meta sequences (such as \".\" and \"..\") from the URL.");

static AddInX xai_inet_canonicalize_url(
	FunctionX(XLL_CSTRINGX, _T("?xll_inet_canonicalize_url"), _T("INET.CANONICALIZE.URL"))
	.Arg(XLL_CSTRINGX, _T("Url"), _T("is the URL to canonicalize."))
	.Arg(XLL_WORDX, _T("Flags"), _T("are optional flags from ICU_*."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a canonicalized Url."))
	.Documentation(_T(""))
);
xcstr WINAPI xll_inet_canonicalize_url(xcstr url, WORD flags)
{
#pragma XLLEXPORT
	static xchar curl[1024];
	DWORD len = 1024;

	try {
		ensure (InternetCanonicalizeUrl(url, curl, &len, flags));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return curl;
}

