#include "field.h"

msado::field::field()
	: cimpl<field,FieldPtr,false>()
{

}

msado::field::field(field const& rhs)
	: cimpl<field, FieldPtr, false>(rhs)
{

}

msado::field::field(field::interface_type* ptr)
	: cimpl<field, FieldPtr, false>(ptr)
{

}

////////////////////////////////////////////////////
// public method

msado::properties msado::field::get_properties() const
{
	properties p;
	(*this)->get_Properties(p.get_address_of());
	return p;
}

long msado::field::get_actual_size() const
{
	long result = 0;
	(*this)->get_ActualSize(&result);
	return result;
}

long msado::field::get_attributes() const
{
	long result = 0;
	(*this)->get_Attributes(&result);
	return result;
}

long msado::field::get_defined_size() const
{
	long result = 0;
	(*this)->get_DefinedSize(&result);
	return result;
}

msado::string msado::field::get_name() const
{
	BSTR bstr = NULL;
	(*this)->get_Name(&bstr);
	return bstr;
}

msado::data_type msado::field::get_type() const
{
	DataTypeEnum dt;
	(*this)->get_Type(&dt);
	return static_cast<data_type>(dt);
}

_variant_t msado::field::get_value() const
{
	VARIANT var;
	VariantInit(&var);

	(*this)->get_Value(&var);
	return _variant_t(var, false);
}

void msado::field::set_value(const _variant_t & pvar)
{
	(*this)->put_Value(pvar);
}

unsigned char msado::field::get_precision()
{
	BYTE pc;
	(*this)->get_Precision(&pc);
	return pc;
}

unsigned char msado::field::get_numeric_scale()
{
	BYTE ns;
	(*this)->get_NumericScale(&ns);
	return ns;
}

bool msado::field::append_chunk(const _variant_t & Data)
{
	HRESULT hr = (*this)->AppendChunk(Data);
	return hr == S_OK;
}

_variant_t msado::field::get_chunk(long Length)
{
	VARIANT var;
	VariantInit(&var);

	(*this)->GetChunk(Length,&var);
	return _variant_t(var, false);
}
_variant_t msado::field::get_original_value() const
{
	VARIANT var;
	VariantInit(&var);

	(*this)->get_OriginalValue(&var);
	return _variant_t(var, false);
}

_variant_t msado::field::get_underlying_value() const
{
	VARIANT var;
	VariantInit(&var);

	(*this)->get_UnderlyingValue(&var);
	return _variant_t(var, false);
}

IUnknownPtr msado::field::get_data_format() const
{
	IUnknownPtr puk;
	(*this)->get_DataFormat(&puk);
	return puk;
}

void msado::field::set_ref_data_format(IUnknown * ppiDF)
{
	(*this)->putref_DataFormat(ppiDF);
}

void msado::field::set_precision(unsigned char pbPrecision)
{
	(*this)->put_Precision(pbPrecision);
}

void msado::field::set_numeric_scale(unsigned char pbNumericScale)
{
	(*this)->put_NumericScale(pbNumericScale);
}

void msado::field::set_type(data_type pDataType)
{
	(*this)->put_Type(static_cast<DataTypeEnum>(pDataType));
}

void msado::field::set_defined_size(long pl)
{
	(*this)->put_DefinedSize(pl);
}

void msado::field::set_attributes(long pl)
{
	(*this)->put_Attributes(pl);
}

long msado::field::get_status() const
{
	long result = 0;
	(*this)->get_Status(&result);
	return result;
}
