#pragma once

#include "msado.h"
#include "property.h"

namespace msado
{
	class properties : public detail::cimpl<properties, PropertiesPtr, false>
	{
	public:
		properties();
		properties(properties const&);
		properties(interface_type*);
	public:
		long get_count() const;
		IUnknownPtr _new_enum();
		bool refresh();

		property get_item(const _variant_t & Index) const;
	};
}
