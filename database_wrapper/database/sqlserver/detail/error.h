#pragma once

#include "msado.h"

namespace msado
{
	class error
		: public detail::cimpl<error, ErrorPtr, false>
	{
	public:
		error();
		error(error const&);
		error(interface_type* ptr);
	public:
		long get_number() const;
		string get_source() const;
		string get_description() const;
		string get_help_file() const;
		long get_help_context() const;
		string get_sql_state() const;
		long get_native_error() const;
	};
}
