#include "error.h"

/////////////////////////////////////////////
// ctor method

msado::error::error()
	: cimpl<error,ErrorPtr,false>()
{

}

msado::error::error(error const& rhs)
	: cimpl<error, ErrorPtr, false>(rhs)
{

}

msado::error::error(interface_type* ptr)
	: cimpl<error, ErrorPtr, false>(ptr)
{

}

/////////////////////////////////////////////////
// public method

long msado::error::get_number() const
{
	long result;
	(*this)->get_Number(&result);
	return result;
}

msado::string msado::error::get_source() const
{
	BSTR bstr = NULL;
	(*this)->get_Source(&bstr);
	return bstr;
}

msado::string msado::error::get_description() const
{
	BSTR bstr = NULL;
	(*this)->get_Description(&bstr);
	return bstr;
}

msado::string msado::error::get_help_file() const
{
	BSTR bstr = NULL;
	(*this)->get_HelpFile(&bstr);
	return bstr;
}

long msado::error::get_help_context() const
{
	long result=0;
	(*this)->get_HelpContext(&result);
	return result;
}

msado::string msado::error::get_sql_state() const
{
	BSTR bstr = NULL;
	(*this)->get_SQLState(&bstr);
	return bstr;
}

long msado::error::get_native_error() const
{
	long result = 0;
	(*this)->get_NativeError(&result);
	return result;
}
