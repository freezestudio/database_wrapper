#pragma once

#include "msado.h"

namespace msado
{
	class property
		: public detail::cimpl<property,PropertyPtr,false>
	{
	public:
		property();
		property(property const&);
		property(interface_type*);
	public:
		_variant_t get_value() const;
		void set_value(const _variant_t & pval);
		string get_name() const;
		data_type get_type() const;
		long get_attributes() const;
		void set_attributes(long plAttributes);
	};
}
