#include "command.h"

///////////////////////////////////////////
// ctor dtor internal method

msado::command::command()
	: cimpl<command,_CommandPtr>()
{
}

msado::command::command(command const& rhs)
	: cimpl<command, _CommandPtr>(rhs)
{
}

msado::command::command(interface_type* ptr)
	: cimpl<command, _CommandPtr>(ptr)
{
}

msado::command::command(CLSID const& clsid)
	: cimpl<command, _CommandPtr>(clsid)
{
}

msado::command::~command()
{
}

bool msado::command::enable_event(command_event* pevent,bool enabled/* = true*/)
{
	return pevent->enable_event(enabled);
}

//////////////////////////////////////////////////
// public method

msado::properties msado::command::get_properties() const
{
	properties p;
	ptr_->get_Properties(p.get_address_of());
	return p;
}

msado::connection msado::command::get_active_connection() const
{
	connection conn;
	ptr_->get_ActiveConnection(conn.get_address_of());
	return conn;
}

bool msado::command::set_ref_active_connection(msado::connection& conn)
{
	HRESULT hr=ptr_->putref_CommandStream(static_cast<IUnknown*>(
		static_cast<IDispatch*>(conn.get())));
	return hr == S_OK;
}

bool msado::command::set_active_connection(_variant_t const& conn)
{
	HRESULT hr=ptr_->put_ActiveConnection(conn);
	return hr == S_OK;
}

msado::string msado::command::get_command_text() const
{
	BSTR bstr = NULL;
	ptr_->get_CommandText(&bstr);
	return bstr;
}

void msado::command::set_command_text(string const& pbstr)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstr);
	ptr_->put_CommandText(bstr.GetBSTR());	
}

long msado::command::get_command_timeout() const
{
	long result = 0;
	ptr_->get_CommandTimeout(&result);
	return result;
}

void msado::command::set_command_timeout(long pl/* = 30second*/)
{
	ptr_->put_CommandTimeout(pl);
}

bool msado::command::get_prepared() const
{
	VARIANT_BOOL vb = 0;
	ptr_->get_Prepared(&vb);
	return vb == VARIANT_TRUE;
}

void msado::command::set_prepared(bool pfPrepared)
{
	ptr_->put_Prepared(pfPrepared);
}

//TODO
msado::recordset msado::command::execute(
	long * RecordsAffected,
	string const& Parameters,
	long Options/* = CommandTypeEnum::adCmdText*/)
{
	recordset set;

	_bstr_t param = detail::string_to_bstr_t(Parameters);
	ptr_->Execute(&_variant_t(*RecordsAffected), &_variant_t(param), Options, set.get_address_of());
	return set;
}

msado::parameter msado::command::create_parameter(
	string const& Name,
	data_type Type,
	parameter_direction Direction,
	long Size,
	const _variant_t & Value/* = vtMissing*/)
{
	parameter p;

	_bstr_t name = detail::string_to_bstr_t(Name);
	ptr_->CreateParameter(name.GetBSTR(),
		static_cast<DataTypeEnum>(Type),
		static_cast<ParameterDirectionEnum>(Direction),
		Size,
		Value,
		p.get_address_of());
	return p;
}

msado::parameters msado::command::get_parameters() const
{
	parameters p;
	ptr_->get_Parameters(p.get_address_of());
	return p;
}

void msado::command::set_command_type(command_type plCmdType)
{
	ptr_->put_CommandType(static_cast<CommandTypeEnum>(plCmdType));
}

msado::command_type msado::command::get_command_type() const
{
	CommandTypeEnum ct;
	ptr_->get_CommandType(&ct);
	return static_cast<command_type>(ct);
}

msado::string msado::command::get_name() const
{
	BSTR bstr = NULL;
	ptr_->get_Name(&bstr);
	return bstr;
}

void msado::command::set_name(string const& pbstrName)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstrName);
	ptr_->put_Name(bstr.GetBSTR());
}

long msado::command::get_state() const
{
	long result = 0;
	ptr_->get_State(&result);
	return result;
}

bool msado::command::cancel()
{
	HRESULT hr=ptr_->Cancel();
	return hr == S_OK;
}

void msado::command::set_ref_command_stream(IUnknown * pvStream)
{
	ptr_->putref_CommandStream(pvStream);
}

_variant_t msado::command::get_command_stream() const
{
	VARIANT var;
	VariantInit(&var);

	HRESULT hr=ptr_->get_CommandStream(&var);
	return _variant_t(var, false);
}

void msado::command::set_dialect(string const& pbstrDialect)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstrDialect);
	ptr_->put_Dialect(bstr.GetBSTR());
}

msado::string msado::command::get_dialect() const
{
	BSTR bstr = NULL;
	ptr_->get_Dialect(&bstr);
	return bstr;
}

void msado::command::set_named_parameters(bool pfNamedParameters)
{
	ptr_->put_NamedParameters(pfNamedParameters ? VARIANT_TRUE : VARIANT_FALSE);
}

bool msado::command::get_named_parameters() const
{
	VARIANT_BOOL vb = 0;
	HRESULT hr = ptr_->get_NamedParameters(&vb);
	return vb == VARIANT_TRUE;
}

///////////////////////////////////////////////
// event method

void msado::command::on_begin_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::command::on_commit_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::command::on_rollback_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::command::on_executing(string const& source,
	cursor_type cursortype,
	lock_type locktype,
	long options,
	command const& cmd,
	recordset const& set,
	connection const& conn)
{

}

void msado::command::on_executed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::command::on_connecting(
	string const& connstr, string const& user, string const& password, long options)
{

}

void msado::command::on_connected(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::command::on_disconnected()
{

}

