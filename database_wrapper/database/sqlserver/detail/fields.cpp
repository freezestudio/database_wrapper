#include "fields.h"

msado::fields::fields()
	:cimpl<fields,FieldsPtr,false>()
{
	
}

msado::fields::fields(fields const& rhs)
	: cimpl<fields, FieldsPtr,false>(rhs)
{

}

msado::fields::fields(fields::interface_type* ptr)
	: cimpl<fields, FieldsPtr, false>(ptr)
{

}

///////////////////////////////////////////////
// public method

long msado::fields::get_count() const
{
	long result = 0;
	(*this)->get_Count(&result);
	return result;
}

IUnknownPtr msado::fields::_new_enum()
{
	IUnknownPtr puk;
	(*this)->_NewEnum(&puk);
	return puk;
}

bool msado::fields::refresh()
{
	HRESULT hr = (*this)->Refresh();
	return S_OK;
}

msado::field msado::fields::get_item(const _variant_t & Index) const
{
	field f;
	(*this)->get_Item(Index, f.get_address_of());
	return f;
}

bool msado::fields::_delete(const _variant_t & Index)
{
	HRESULT hr = (*this)->Delete(Index);
	return S_OK;
}

bool msado::fields::append(string const& Name,
	data_type Type,
	long DefinedSize,
	field_attribute Attrib,
	const _variant_t & FieldValue/* = vtMissing*/)
{
	_bstr_t name = detail::string_to_bstr_t(Name);
	HRESULT hr = (*this)->Append(name.GetBSTR(), 
		static_cast<DataTypeEnum>(Type),
		DefinedSize,
		static_cast<FieldAttributeEnum>(Attrib),
		FieldValue);
	return S_OK;
}

bool msado::fields::update()
{
	HRESULT hr = (*this)->Update();
	return S_OK;
}

bool msado::fields::resync(_resync ResyncValues)
{
	HRESULT hr = (*this)->Resync(static_cast<ResyncEnum>(ResyncValues));
	return S_OK;
}

bool msado::fields::cancel_update()
{
	HRESULT hr=(*this)->CancelUpdate();
	return S_OK;
}