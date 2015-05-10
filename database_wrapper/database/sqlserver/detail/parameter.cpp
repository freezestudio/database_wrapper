#include "parameter.h"

/////////////////////////////////////////
// ctor method

msado::parameter::parameter()
	: cimpl<parameter,_ParameterPtr>()
{

}

msado::parameter::parameter(parameter const& rhs)
	: cimpl<parameter, _ParameterPtr>(rhs)
{

}

msado::parameter::parameter(interface_type* ptr)
	: cimpl<parameter, _ParameterPtr>(ptr)
{

}

msado::parameter::parameter(CLSID const& clsid)
	: cimpl<parameter, _ParameterPtr>(clsid)
{

}

////////////////////////////////////////////////
// public method

msado::properties msado::parameter::get_properties() const
{
	properties p;
	ptr_->get_Properties(p.get_address_of());
	return p;
}

msado::string msado::parameter::get_name() const
{
	BSTR bstr = NULL;
	ptr_->get_Name(&bstr);
	return bstr;
}

void msado::parameter::set_name(string const& pbstr)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstr);
	ptr_->put_Name(bstr.GetBSTR());
}

_variant_t msado::parameter::get_value()
{
	VARIANT var;
	VariantInit(&var);

	ptr_->get_Value(&var);
	return _variant_t(var, false);
}

void msado::parameter::set_value(const _variant_t & pvar)
{
	ptr_->put_Value(pvar);
}

msado::data_type msado::parameter::get_type() const
{
	DataTypeEnum dt;
	ptr_->get_Type(&dt);
	return static_cast<data_type>(dt);
}

void msado::parameter::set_type(data_type psDataType)
{
	ptr_->put_Type(static_cast<DataTypeEnum>(psDataType));
}

void msado::parameter::set_direction(parameter_direction plParmDirection)
{
	ptr_->put_Direction(static_cast<ParameterDirectionEnum>(plParmDirection));
}

msado::parameter_direction msado::parameter::get_direction()
{
	ParameterDirectionEnum pd;
	ptr_->get_Direction(&pd);
	return static_cast<parameter_direction>(pd);
}

void msado::parameter::set_precision(unsigned char pbPrecision)
{
	ptr_->put_Precision(pbPrecision);
}

unsigned char msado::parameter::get_precision()
{
	unsigned char pb;
	ptr_->get_Precision(&pb);
	return pb;
}

void msado::parameter::set_numericScale(unsigned char pbScale)
{
	ptr_->put_NumericScale(pbScale);
}

unsigned char msado::parameter::get_numericScale()
{
	unsigned char ns;
	ptr_->get_NumericScale(&ns);
	return ns;
}

void msado::parameter::set_size(long pl)
{
	ptr_->put_Size(pl);
}

long msado::parameter::get_size()
{
	long result = 0;
	ptr_->get_Size(&result);
	return result;
}

bool msado::parameter::append_chunk(const _variant_t & Val)
{
	HRESULT hr = ptr_->AppendChunk(Val);
	return hr == S_OK;
}

long msado::parameter::get_attributes()
{
	long result = 0;
	ptr_->get_Attributes(&result);
	return result;
}

void msado::parameter::set_attributes(long plParmAttribs)
{
	ptr_->put_Attributes(plParmAttribs);
}

