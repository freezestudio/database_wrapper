#pragma once

#include "msado.h"
#include "properties.h"

namespace msado
{
	class parameter
		: public detail::cimpl<parameter, _ParameterPtr>
	{
	public:
		parameter();
		parameter(parameter const&);
		parameter(interface_type* ptr);
		explicit parameter(CLSID const& clsid);
	public:
		properties get_properties() const;

		string get_name() const;
		void set_name(string const& pbstr);
		_variant_t get_value();
		void set_value(const _variant_t & pvar);
		data_type get_type() const;
		void set_type(data_type psDataType);
		void set_direction(parameter_direction plParmDirection);
		parameter_direction get_direction();
		void set_precision(unsigned char pbPrecision);
		unsigned char get_precision();
		void set_numericScale(unsigned char pbScale);
		unsigned char get_numericScale();
		void set_size(long pl);
		long get_size();
		bool append_chunk(const _variant_t & Val);
		long get_attributes();
		void set_attributes(long plParmAttribs);
	};
}