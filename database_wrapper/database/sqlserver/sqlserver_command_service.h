#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的命令服务。
	//
	template<typename DataBase>
	class sqlserver_command_service
		: public detail::basic_service<sqlserver_command_service<DataBase> >
	{
	public:
		typedef msado::command implement_type;
		typedef DataBase database_type;
	public:
		sqlserver_command_service()
			:database_()
		{

		}

		explicit sqlserver_command_service(database_type const& db)
			:database_(db)
		{

		}

		sqlserver_command_service(sqlserver_command_service const& rhs)
			: database_(rhs.database_)
		{

		}
	//public:
	//	database_type get_database() const
	//	{
	//		return database_;
	//	}
	public:
		bool open(implement_type& impl);
		bool close(implement_type& impl);
	public:
		sqlserver_connection<DataBase> get_active_connection() const;
		bool set_ref_active_connection(sqlserver_connection<DataBase>& conn);
		bool set_active_connection(string const& connstr);
		msado::string get_command_text() const;
		void set_command_text(string const& pbstr);
		long get_command_timeout() const;
		void set_command_timeout(long pl = 30/*second*/);
		bool get_prepared();
		void set_prepared(bool pfPrepared);
		sqlserver_recordset<DataBase> execute(
			long * recordsaffected,
			string const& parameters,
			long options = msado::/*command_type::*/cmd_text);

		//parameter create_parameter(
		//	string const& name,
		//	data_type type,
		//	parameter_direction direction,
		//	long size,
		//	const _variant_t & Value = vtMissing);

		parameter create_parameter(
			string const& name,	data_type type,
			parameter_direction direction,long size,
			long value);

		parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			unsigned long value);

		parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			double value);

		parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			string const& value);

		parameter create_parameter(
			string const& name, data_type type,
			parameter_direction direction, long size,
			datetime const& value);

		parameters get_parameters() const;
		void set_command_type(command_type plCmdType);
		command_type get_command_type() const;
		msado::string get_name() const;
		void set_name(string const& pbstrName);
		long get_state() const;
		bool cancel();
		//void set_ref_command_stream(IUnknown * pvStream);
		//_variant_t get_command_stream() const;
		//void set_dialect(string const& pbstrDialect);
		//msado::string get_dialect() const;
		//void set_named_parameters(bool pfNamedParameters);
		//bool get_named_parameters() const;
	private:
		bool opened(implement_type const& impl) const;
	private:
		database_type database_;
	};
}
