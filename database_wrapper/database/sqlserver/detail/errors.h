#pragma once

#include "msado.h"
#include "error.h"

namespace msado
{
	class errors
		: public detail::cimpl<errors, ErrorsPtr, false>
	{
	public:
		errors();
		errors(errors const&);
		errors(interface_type*);
	public:
		long get_count() const;
		IUnknownPtr _new_enum();
		bool refresh();
		error get_item(	const _variant_t & Index) const;
		bool clear();
	};
}
