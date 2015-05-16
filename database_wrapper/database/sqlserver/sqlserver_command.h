#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的命令操作
	//
	template<typename DataBase, 
		template<typename>class SqlService/* = sqlserver_command_service*/>
	class sqlserver_command
		: public detail::basic_command<DataBase, SqlService>
	{
	public:
		explicit sqlserver_command(
			CLSID const& clsid=msado::Clsid::ADODB_Command())
			: detail::basic_command<DataBase, SqlService>(clsid)
		{
		}

		sqlserver_command(DataBase const& db, msado::command const& cmd)
			: detail::basic_command<DataBase,SqlService>(db,cmd)
		{

		}
	public:
		sqlserver_connect<DataBase> get_active_connection() const;
		bool set_ref_active_connection(sqlserver_connection<DataBase>& conn);
		bool set_active_connection(string const& connstr);
		msado::string get_command_text() const;
		void set_command_text(string const& pbstr);
		long get_command_timeout() const;
		void set_command_timeout(long pl = 30/*second*/);
		bool get_prepared() const;
		void set_prepared(bool pfPrepared);
		sqlserver_recordset<DataBase> execute(
			long * recordsaffected,
			string const& parameters,
			long options = msado::/*command_type::*/cmd_text);
		msado::parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			long value);

		msado::parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			unsigned long value);

		msado::parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			double value);

		msado::parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			string const& value);

		msado::parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			datetime const& value);

		msado::parameters get_parameters() const;
		void set_command_type(command_type plCmdType);
		msado::command_type get_command_type() const;
		msado::string get_name() const;
		void set_name(string const& pbstrName);
		long get_state() const;
		bool cancel();
	public:
		typename implement_type::interface_type const* get_internal_ptr() const;
		typename implement_type::interface_type * get_internal_ptr();
	};
}

namespace database
{
	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_connect<DataBase> 
		sqlserver_command<DataBase, SqlService>::get_active_connection() const
	{
		return get_service().get_active_connection(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_command<DataBase, SqlService>::set_ref_active_connection(
		sqlserver_connection<DataBase>& conn)
	{
		return get_service().set_ref_active_connection(get_implement(),conn);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_command<DataBase, SqlService>::set_active_connection(
		string const& connstr)
	{
		return get_service().set_active_connection(get_implement(),connstr);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::string sqlserver_command<DataBase, SqlService>::get_command_text(
		) const
	{
		return get_service().get_command_text(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_command<DataBase, SqlService>::set_command_text(
		string const& pbstr)
	{
		return get_service().set_command_text(get_implement(),pbstr);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_command<DataBase, SqlService>::get_command_timeout(
		) const
	{
		return get_service().get_command_timeout(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_command<DataBase, SqlService>::set_command_timeout(
		long pl/* = 30second*/)
	{
		return get_service().set_command_timeout(get_implement(),pl);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_command<DataBase, SqlService>::get_prepared() const
	{
		return get_service().get_prepared(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_command<DataBase, SqlService>::set_prepared(
		bool pfPrepared)
	{
		return get_service().set_prepared(get_implement(),pfPrepared);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_recordset<DataBase> 
		sqlserver_command<DataBase, SqlService>::execute(
		long * recordsaffected,
		string const& parameters,
		long options/* = msado::command_type::cmd_text*/)
	{
		return get_service().execute(get_implement(),
			recordsaffected,parameters,options);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameter 
		sqlserver_command<DataBase, SqlService>::create_parameter(
		string const& name, data_type type,
		parameter_direction direction, long size,
		long value)
	{
		return get_service().create_parameter(get_implement(),
			name, type, direction, size, value);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameter 
		sqlserver_command<DataBase, SqlService>::create_parameter(
		string const& name, data_type type,
		parameter_direction direction, long size,
		unsigned long value)
	{
		return get_service().create_parameter(get_implement(),
			name, type, direction, size, value);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameter 
		sqlserver_command<DataBase, SqlService>::create_parameter(
		string const& name, data_type type,
		parameter_direction direction, long size,
		double value)
	{
		return get_service().create_parameter(get_implement(),
			name, type, direction, size, value);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameter 
		sqlserver_command<DataBase, SqlService>::create_parameter(
		string const& name, data_type type,
		parameter_direction direction, long size,
		string const& value)
	{
		return get_service().create_parameter(get_implement(),
			name, type, direction, size, value);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameter 
		sqlserver_command<DataBase, SqlService>::create_parameter(
		string const& name, data_type type,
		parameter_direction direction, long size,
		datetime const& value)
	{
		return get_service().create_parameter(get_implement(),
			name,type,direction,size,value);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::parameters 
		sqlserver_command<DataBase, SqlService>::get_parameters() const
	{
		return get_service().get_parameters(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_command<DataBase, SqlService>::set_command_type(
		command_type plCmdType)
	{
		get_service().set_command_type(get_implement(),plCmdType);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::command_type 
		sqlserver_command<DataBase, SqlService>::get_command_type() const
	{
		return get_service().get_command_type(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::string sqlserver_command<DataBase, SqlService>::get_name() const
	{
		return get_service().get_name(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_command<DataBase, SqlService>::set_name(
		string const& pbstrName)
	{
		get_service().set_name(get_implement(),pbstrName);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_command<DataBase, SqlService>::get_state() const
	{
		return get_service().get_state(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_command<DataBase, SqlService>::cancel()
	{
		return get_service().cancel(get_implement());
	}

	template<typename DataBase,template<typename>class SqlService>
	inline typename sqlserver_command<DataBase,SqlService>::
		implement_type::interface_type const*
		sqlserver_command<DataBase, SqlService>::get_internal_ptr() const
	{
		return get_implement().get();
	}

	template<typename DataBase, template<typename>class SqlService>
	inline typename sqlserver_command<DataBase, SqlService>::
		implement_type::interface_type *
		sqlserver_command<DataBase, SqlService>::get_internal_ptr()
	{
		return get_implement().get();
	}
}