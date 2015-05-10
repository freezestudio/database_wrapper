#include "property.h"

msado::property::property()
	: cimpl<property,PropertyPtr,false>()
{

}

msado::property::property(property const& rhs)
	: cimpl<property, PropertyPtr, false>(rhs)
{

}

msado::property::property(interface_type* ptr)
	: cimpl<property, PropertyPtr, false>(ptr)
{

}

///////////////////////////////////////////
// public method
_variant_t msado::property::get_value() const
{
	VARIANT var;
	VariantInit(&var);
	(*this)->get_Value(&var);
	return _variant_t(var, false);
}

void msado::property::set_value(const _variant_t & pval)
{
	(*this)->put_Value(pval);
}

msado::string msado::property::get_name() const
{
	BSTR bstr = NULL;
	(*this)->get_Name(&bstr);
	return bstr;
}

msado::data_type msado::property::get_type() const
{
	DataTypeEnum dt;
	(*this)->get_Type(&dt);
	return static_cast<data_type>(dt);
}

long msado::property::get_attributes() const
{
	long result = 0;
	(*this)->get_Attributes(&result);
	return result;
}

void msado::property::set_attributes(long plAttributes)
{
	(*this)->put_Attributes(plAttributes);
}
