#pragma once

#include "msado.h"
#include "field.h"

namespace msado
{
	class fields
		: public detail::cimpl<fields, FieldsPtr, false>
	{
	public:
		fields();
		fields(fields const&);
		fields(interface_type*);
	public:
		long get_count() const;
		IUnknownPtr _new_enum();
		bool refresh();

		field get_item(const _variant_t & Index) const;
		field get_item(const _variant_t & Index);

		//bool _append(string const& Name,
		//	data_type Type,
		//	long DefinedSize,
		//	field_attribute Attrib);
		bool _delete(const _variant_t & Index);

		bool append(string const& Name,
			data_type Type,
			long DefinedSize,
			field_attribute Attrib,
			const _variant_t & FieldValue = vtMissing);
		bool update();
		bool resync(_resync ResyncValues);
		bool cancel_update();
	};
}
