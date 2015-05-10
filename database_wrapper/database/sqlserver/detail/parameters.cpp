#include "parameters.h"

msado::parameters::parameters()
	: cimpl<parameters, ParametersPtr,false>()
{

}

msado::parameters::parameters(parameters const& rhs)
	: cimpl<parameters, ParametersPtr,false>(rhs)
{

}

msado::parameters::parameters(interface_type* ptr)
	: cimpl<parameters, ParametersPtr, false>(ptr)
{

}

///////////////////////////////////////////////////
// public method

long msado::parameters::get_count() const
{
	long result = 0;
	(*this)->get_Count(&result);
	return result;
}

IUnknownPtr msado::parameters::_new_enum()
{
	IUnknownPtr puk;
	(*this)->_NewEnum(&puk);
	return puk;
}

bool msado::parameters::refresh()
{
	HRESULT hr = (*this)->Refresh();
	return hr == S_OK;
}


bool msado::parameters::append(IDispatch * Object)
{
	HRESULT hr = (*this)->Append(Object);
	return hr == S_OK;
}

bool msado::parameters::_delete(const _variant_t & Index)
{
	HRESULT hr = (*this)->Delete(Index);
	return hr == S_OK;
}


msado::parameter msado::parameters::get_item(const _variant_t & Index) const
{
	parameter p;
	(*this)->get_Item(Index, p.get_address_of());
	return p;
}
