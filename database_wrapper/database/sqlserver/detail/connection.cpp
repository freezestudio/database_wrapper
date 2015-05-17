//
//

#include "connection.h"

//////////////////////////////////////////////////////
// ctor -- dtor internal methord

msado::connection::connection()
	: cimpl<connection, _ConnectionPtr>()
{
}

msado::connection::connection(connection const& rhs)
	: cimpl<connection, _ConnectionPtr>(rhs)
{
}

msado::connection::connection(connection::interface_type* ptr)
	: cimpl<connection, _ConnectionPtr>(ptr)
{	
}

msado::connection::connection(CLSID const& clsid)
	: cimpl<connection, _ConnectionPtr>(clsid)
{
}

msado::connection::~connection()
{
}

bool msado::connection::enable_event(connection_event* pevent, bool enabled/* = true*/)
{
	return pevent->enable_event(enabled);
}

//////////////////////////////////////////////////////////////
// public method

msado::properties msado::connection::get_properties() const
{
	properties result;
	HRESULT hr = (*this)->get_Properties(result.get_address_of());
	return result;
}

msado::string msado::connection::get_connection_string() const
{
	msado::string str;

	BSTR bstr = NULL;
	HRESULT hr = (*this)->get_ConnectionString(&bstr);
	if (SUCCEEDED(hr))
	{
		//_bstr_t result(bstr, false);
		//str = detail::_bstr_t_to_string(result);
		return bstr;
	}

	return str;
}

void msado::connection::set_connection_string(msado::string const& pbstr)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstr);
	HRESULT hr = (*this)->put_ConnectionString(bstr.GetBSTR());
}

long msado::connection::get_command_timeout() const
{
	long timeout = 0;
	(*this)->get_CommandTimeout(&timeout);
	return timeout;
}

void msado::connection::set_command_timeout(long plTimeout/* = 30 seconds*/)
{
	(*this)->put_CommandTimeout(plTimeout);
}

long msado::connection::get_connection_timeout() const
{
	long timeout = 0;
	(*this)->get_ConnectionTimeout(&timeout);
	return timeout;
}

void msado::connection::set_connection_timeout(long plTimeout /*= 15 seconds*/)
{
	(*this)->put_ConnectionTimeout(plTimeout);
}

msado::string msado::connection::get_version() const
{
	BSTR bstr = NULL;
	(*this)->get_Version(&bstr);
	return detail::_bstr_t_to_string(_bstr_t(bstr, false));
}

bool msado::connection::close()
{
	HRESULT hr = (*this)->Close();
	return hr == S_OK;
}

msado::recordset msado::connection::execute(
	string const& CommandText,
	long& RecordsAffected,
	command_type Options/* = cmd_text*/)
{
	recordset set(Clsid::ADODB_Recordset());
	VARIANT var;
	VariantInit(&var);
	(*this)->Execute(
		detail::string_to_bstr_t(CommandText).GetBSTR(),
		&var,
		Options,
		set.get_address_of());
	RecordsAffected = var.lVal;

	return set;
}

long msado::connection::begin_trans()
{
	long result = 0;
	(*this)->BeginTrans(&result);
	return 0;
}

bool msado::connection::commit_trans()
{
	HRESULT hr = (*this)->CommitTrans();
	return hr == S_OK;
}

bool msado::connection::rollback_trans()
{
	HRESULT hr = (*this)->RollbackTrans();
	return hr == S_OK;
}

bool msado::connection::open(
	string const& ConnectionString,
	string const& UserID,
	string const& Password,
	connect_option Options/* = connect_unspecified*/)
{
	_bstr_t conn = detail::string_to_bstr_t(ConnectionString);
	_bstr_t user = detail::string_to_bstr_t(UserID);
	_bstr_t pwd = detail::string_to_bstr_t(Password);

	HRESULT hr = (*this)->Open(
		conn.GetBSTR(),
		user.GetBSTR(),
		pwd.GetBSTR(),
		static_cast<long>(Options));
	return hr == S_OK;
}

msado::errors msado::connection::get_errors() const
{
	errors e;
	(*this)->get_Errors(e.get_address_of());
	return e;
}

msado::string msado::connection::get_default_database() const
{
	BSTR bstr = NULL;
	(*this)->get_DefaultDatabase(&bstr);
	return detail::_bstr_t_to_string(_bstr_t(bstr, false));
}

void msado::connection::set_default_database(string const& pbstr)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstr);
	(*this)->put_DefaultDatabase(bstr.GetBSTR());
}

msado::isolation_level msado::connection::get_isolation_level() const
{
	long level = 0;
	(*this)->get_IsolationLevel(
		reinterpret_cast<IsolationLevelEnum*>(&level));
	return static_cast<isolation_level>(level);
}

void msado::connection::set_isolation_level(isolation_level Level)
{
	(*this)->put_IsolationLevel(
		static_cast<IsolationLevelEnum>(Level));
}

long msado::connection::get_attributes() const
{
	long result = 0;
	(*this)->get_Attributes(&result);
	return result;
}

void msado::connection::set_attributes(long plAttr)
{
	(*this)->put_Attributes(plAttr);
}

msado::cursor_location msado::connection::get_cursor_location() const
{
	long result = 0;
	(*this)->get_CursorLocation(
		reinterpret_cast<CursorLocationEnum*>(&result));
	return static_cast<cursor_location>(result);
}

void msado::connection::set_cursor_location(cursor_location plCursorLoc)
{
	(*this)->put_CursorLocation(
		static_cast<CursorLocationEnum>(plCursorLoc));
}

msado::connect_mode msado::connection::get_mode() const
{
	long result = 0;
	(*this)->get_Mode(reinterpret_cast<ConnectModeEnum*>(&result));
	return static_cast<connect_mode>(result);
}

void msado::connection::set_mode(connect_mode plMode)
{
	(*this)->put_Mode(static_cast<ConnectModeEnum>(plMode));
}

msado::string msado::connection::get_provider() const
{
	BSTR bstr = NULL;
	(*this)->get_Provider(&bstr);
	return detail::_bstr_t_to_string(_bstr_t(bstr, false));
}

void msado::connection::set_provider(string const& pbstr)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstr);
	(*this)->put_Provider(bstr.GetBSTR());
}

long msado::connection::get_state() const
{
	long result = 0;
	ATLASSERT(this != NULL);
	if (this != NULL)
		(*this)->get_State(&result);
	return result;
}

msado::recordset msado::connection::open_schema(
	schema Schema,
	const _variant_t & Restrictions/* = vtMissing*/,
	const _variant_t & SchemaID/* = vtMissing*/)
{
	recordset set(Clsid::ADODB_Recordset());
	(*this)->OpenSchema(
		static_cast<SchemaEnum>(Schema),
		Restrictions, SchemaID, set.get_address_of());
	return set;
}

bool msado::connection::cancel()
{
	HRESULT hr = (*this)->Cancel();
	return hr == S_OK;
}

///////////////////////////////////////////////
// event method

void msado::connection::on_begin_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::connection::on_commit_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::connection::on_rollback_transed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::connection::on_executing(string const& source,
	cursor_type cursortype,
	lock_type locktype,
	long options,
	command const& cmd,
	recordset const& set,
	connection const& conn)
{

}

void msado::connection::on_executed(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::connection::on_connecting(
	string const& connstr, string const& user, string const& password, long options)
{

}

void msado::connection::on_connected(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::connection::on_disconnected()
{

}

