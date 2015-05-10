#pragma once

#include "msado.h"
#include "properties.h"

namespace msado
{
	class field 
		: public detail::cimpl<field, FieldPtr, false>
	{
	public:
		field();
		field(field const&);
		field(interface_type* ptr);
	public:
		properties get_properties() const;
		long get_actual_size() const;
		long get_attributes() const;
		long get_defined_size() const;
		string get_name() const;
		data_type get_type() const;
		_variant_t get_value() const;
		void set_value(const _variant_t & pvar);
		unsigned char get_precision();
		unsigned char get_numeric_scale();
		bool append_chunk(const _variant_t & Data);
		_variant_t get_chunk(long Length);
		_variant_t get_original_value() const;
		_variant_t get_underlying_value() const;
		IUnknownPtr get_data_format() const;
		void set_ref_data_format(IUnknown * ppiDF);
		void set_precision(unsigned char pbPrecision);
		void set_numeric_scale(unsigned char pbNumericScale);
		void set_type(data_type pDataType);
		void set_defined_size(long pl);
		void set_attributes(long pl);
		long get_status() const;
	};
}