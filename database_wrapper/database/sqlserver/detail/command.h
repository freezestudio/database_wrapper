#pragma once

#include "msado.h"
#include "properties.h"
#include "connection.h"
#include "parameter.h"
#include "parameters.h"
#include "event_command.h"

namespace msado
{
	class command 
		: public detail::cimpl<command, _CommandPtr>
	{
	public:
		command();
		command(command const&);
		command(interface_type*);
		explicit command(CLSID const& clsid);
		~command();
	public:
		properties get_properties() const;
		connection get_active_connection() const;
		bool set_ref_active_connection(connection& conn);
		bool set_active_connection(_variant_t const& conn);
		msado::string get_command_text() const;
		void set_command_text(string const& pbstr);
		long get_command_timeout() const;
		void set_command_timeout(long pl = 30/*second*/);
		bool get_prepared() const;
		void set_prepared(bool pfPrepared);
		recordset execute(
			long * RecordsAffected,
			string const& Parameters,
			long Options = /*CommandTypeEnum::*/adCmdText);
		parameter create_parameter(
			string const& Name,
			data_type Type,
			parameter_direction Direction,
			long Size,
			const _variant_t & Value = vtMissing);
		parameters get_parameters() const;
		void set_command_type(command_type plCmdType);
		command_type get_command_type() const;
		msado::string get_name() const;
		void set_name(string const& pbstrName);
		long get_state() const;
		bool cancel();
		void set_ref_command_stream(IUnknown * pvStream);
		_variant_t get_command_stream() const;
		void set_dialect(string const& pbstrDialect);
		msado::string get_dialect() const;
		void set_named_parameters(bool pfNamedParameters);
		bool get_named_parameters() const;

		//ÊÂ¼þ
	public:
		bool enable_event(command_event* pevent, bool enabled = true);
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
	};
}
