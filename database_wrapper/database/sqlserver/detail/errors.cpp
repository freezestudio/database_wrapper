#include "errors.h"

msado::errors::errors()
	: cimpl<errors,ErrorsPtr,false>()
{

}

msado::errors::errors(errors const& rhs)
	: cimpl<errors, ErrorsPtr, false>(rhs)
{

}

msado::errors::errors(interface_type* ptr)
	: cimpl<errors, ErrorsPtr, false>(ptr)
{

}

///////////////////////////////////////////////////
// public method

long msado::errors::get_count() const
{
	long result = 0;
	(*this)->get_Count(&result);
	return result;
}

IUnknownPtr msado::errors::_new_enum()
{
	IUnknownPtr puk;
	(*this)->_NewEnum(&puk);
	return puk;
}

bool msado::errors::refresh()
{
	HRESULT hr = (*this)->Refresh();
	return hr == S_OK;
}

msado::error msado::errors::get_item(const _variant_t & Index) const
{
	error e;
	(*this)->get_Item(Index, e.get_address_of());
	return e;
}

bool msado::errors::clear()
{
	HRESULT hr = (*this)->Clear();
	return hr == S_OK;
}
