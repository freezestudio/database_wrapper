#pragma once

#include "msado.h"
#include "parameter.h"

namespace msado
{
	class parameters
		: public detail::cimpl<parameters, ParametersPtr,false>
	{
	public:
		parameters();
		parameters(parameters const&);
		parameters(interface_type* ptr);
	public:
		long get_count() const;
		IUnknownPtr _new_enum();
		bool refresh();

		bool append(IDispatch * Object);
		bool _delete(const _variant_t & Index);

		parameter get_item(const _variant_t & Index) const;
	};
}