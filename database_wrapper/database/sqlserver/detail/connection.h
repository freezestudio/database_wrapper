#pragma once

#include "msado.h"
#include "properties.h"
#include "recordset.h"
#include "errors.h"

namespace msado
{
	class connection
		: public detail::cimpl<connection, _ConnectionPtr>
	{
	public:
		connection();
		connection(connection const&);
		connection(interface_type* ptr);
		explicit connection(CLSID const& clsid);
		~connection();
	public:
		properties get_properties() const;
		string get_connection_string() const;
		void set_onnectionString(string const& pbstr);
		long get_command_timeout() const;
		void set_command_timeout(long plTimeout = 30/*seconds*/);
		long get_onnection_timeout() const;
		void set_connection_timeout(long plTimeout = 15/*seconds*/);
		string get_version() const;
		bool close();
		recordset execute(string const& CommandText,
			long& RecordsAffected,
			command_type Options = cmd_text);
		long begin_trans();
		bool commit_trans();
		bool rollback_trans();
		bool open(
			string const& ConnectionString,
			string const& UserID,
			string const& Password,
			connect_option Options = connect_unspecified);
		errors get_errors() const;
		string get_default_database() const;
		void set_default_database(string const& pbstr);
		isolation_level get_isolation_level() const;
		void set_isolation_level(isolation_level Level);
		long get_attributes() const;
		void set_attributes(long plAttr);
		cursor_location get_cursor_location() const;
		void set_cursor_location(cursor_location plCursorLoc);
		connect_mode get_mode() const;
		void set_mode(connect_mode plMode);
		string get_provider() const;
		void set_provider(string const& pbstr);
		long get_state() const;
		recordset open_schema(
			schema Schema,
			const _variant_t & Restrictions = vtMissing,
			const _variant_t & SchemaID = vtMissing);
		bool cancel();

		//ÊÂ¼þ
	public:
		void on_begin_transed(error const& e);
		void on_commit_transed(error const& e);
		void on_rollback_transed(error const& e);
		void on_executing(string const&,
			cursor_type,
			lock_type,
			long,
			command const&,
			recordset const&,
			connection const&);
		void on_executed(error const& e);
		void on_connecting(string const&, string const&, string const&, long);
		void on_connected(error const& e);
		void on_disconnected();
	private:
		bool enable_event(bool enable=true);
	private:
		IUnknownPtr eventptr_;
		DWORD dw_event_;
	};
}
