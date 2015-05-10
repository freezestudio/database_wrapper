#include "properties.h"

msado::properties::properties()
	: cimpl<properties,PropertiesPtr,false>()
{

}

msado::properties::properties(properties const& rhs)
	: cimpl<properties, PropertiesPtr, false>(rhs)
{

}

msado::properties::properties(interface_type* ptr)
	: cimpl<properties, PropertiesPtr, false>(ptr)
{

}

//////////////////////////////////////////////////////
// public method


long msado::properties::get_count() const
{
	long result = 0;
	(*this)->get_Count(&result);
	return result;
}

IUnknownPtr msado::properties::_new_enum()
{
	IUnknownPtr puk;
	(*this)->_NewEnum(&puk);
	return puk;
}

bool msado::properties::refresh()
{
	HRESULT hr = (*this)->Refresh();
	return hr == S_OK;
}


msado::property msado::properties::get_item(const _variant_t & Index) const
{
	msado::property p;
	(*this)->get_Item(Index,p.get_address_of());
	return p;
}
