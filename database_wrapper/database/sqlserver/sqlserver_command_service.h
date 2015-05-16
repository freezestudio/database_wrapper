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
	//public:
	//	bool open(implement_type& impl);
	//	bool close(implement_type& impl);
	public:
		sqlserver_connect<DataBase> get_active_connection(
			implement_type const& impl) const;
		bool set_ref_active_connection(
			implement_type& impl,sqlserver_connect<DataBase>& conn);
		bool set_active_connection(implement_type& impl,string const& connstr);
		msado::string get_command_text(implement_type const& impl) const;
		void set_command_text(implement_type& impl, string const& pbstr);
		long get_command_timeout(implement_type const& impl) const;
		void set_command_timeout(implement_type& impl, long pl = 30/*second*/);
		bool get_prepared(implement_type const& impl) const;
		void set_prepared(implement_type& impl, bool pfPrepared);
		sqlserver_recordset<DataBase> execute(implement_type& impl,
			long * recordsaffected,
			string const& parameters,
			long options = msado::/*command_type::*/cmd_text);

		//parameter create_parameter(
		//	string const& name,
		//	data_type type,
		//	parameter_direction direction,
		//	long size,
		//	const _variant_t & Value = vtMissing);

		msado::parameter create_parameter(implement_type& impl,
			string const& name,	msado::data_type type,
			msado::parameter_direction direction,long size,
			long value);

		msado::parameter create_parameter(implement_type& impl,
			string const& name, msado::data_type type,
			msado::parameter_direction direction, long size,
			unsigned long value);

		msado::parameter create_parameter(implement_type& impl,
			string const& name, msado::data_type type,
			msado::parameter_direction direction, long size,
			double value);

		msado::parameter create_parameter(implement_type& impl,
			string const& name, msado::data_type type,
			msado::parameter_direction direction, long size,
			string const& value);

		msado::parameter create_parameter(implement_type& impl,
			string const& name, msado::data_type type,
			msado::parameter_direction direction, long size,
			datetime const& value);

		msado::parameters get_parameters(implement_type const& impl ) const;
		void set_command_type(implement_type& impl, msado::command_type plCmdType);
		msado::command_type get_command_type(implement_type const& impl ) const;
		msado::string get_name(implement_type const& impl) const;
		void set_name(implement_type& impl,string const& pbstrName);
		long get_state(implement_type const& impl) const;
		bool cancel(implement_type& impl);
		//void set_ref_command_stream(implement_type& impl,IUnknown * pvStream);
		//_variant_t get_command_stream(implement_type const& impl) const;
		//void set_dialect(implement_type& impl,string const& pbstrDialect);
		//msado::string get_dialect(implement_type const& impl) const;
		//void set_named_parameters(implement_type& impl,bool pfNamedParameters);
		//bool get_named_parameters(implement_type const& impl) const;
	//private:
	//	bool opened(implement_type const& impl) const;
	private:
		database_type database_;
	};
}

namespace database
{
	template<typename DataBase>
	inline sqlserver_connect<DataBase>
		sqlserver_command_service<DataBase>::get_active_connection(
			implement_type const& impl) const
	{
		msado::connection conn = impl.get_active_connection();
		return sqlserver_connect<DataBase>(database_, conn);
	}

	template<typename DataBase>
	inline bool sqlserver_command_service<DataBase>::set_ref_active_connection(
		implement_type& impl,sqlserver_connect<DataBase>& conn)
	{
		_vraiant_t var(conn.get_internal_ptr());
		return impl.set_active_connection(var);
	}
	
	template<typename DataBase>
	inline bool sqlserver_command_service<DataBase>::set_active_connection(
		implement_type& impl,string const& connstr)
	{
		return impl.set_active_connection(_variant_t(
			static_cast<const wchar_t*>(connstr)));
	}
		
	template<typename DataBase>
	inline msado::string 
		sqlserver_command_service<DataBase>::get_command_text(
			implement_type const& impl) const
	{
		return impl.get_command_text();
	}
	
	template<typename DataBase>
	inline void sqlserver_command_service<DataBase>::set_command_text(
		implement_type& impl,string const& pbstr)
	{
		impl.set_command_text(pbstr);
	}
	
	template<typename DataBase>
	inline long sqlserver_command_service<DataBase>::get_command_timeout(
		implement_type const& impl) const
	{
		return impl.get_command_timeout();
	}
	
	template<typename DataBase>
	inline void sqlserver_command_service<DataBase>::set_command_timeout(
		implement_type& impl,long pl/* = 30second*/)
	{
		impl.set_command_timeout(static_cast<long>(pl));
	}
	
	template<typename DataBase>
	inline bool sqlserver_command_service<DataBase>::get_prepared(
		implement_type const& impl) const
	{
		return impl.get_prepared();
	}
	
	template<typename DataBase>
	inline void sqlserver_command_service<DataBase>::set_prepared(
		implement_type& impl,bool pfPrepared)
	{
		impl.set_prepared(pfPrepared);
	}
	
	template<typename DataBase>
	inline sqlserver_recordset<DataBase> 
		sqlserver_command_service<DataBase>::execute(implement_type& impl,
		long * recordsaffected,string const& parameters,
		long options/* = msado::command_type::cmd_text*/)
	{
		msado::recordset set=impl.execute(recordsaffected, parameters, options);
		return sqlserver_recordset<DataBase>(database_, set);
	}

	//parameter create_parameter(
	//	string const& name,
	//	data_type type,
	//	parameter_direction direction,
	//	long size,
	//	const _variant_t & Value = vtMissing);

	template<typename DataBase>
	inline msado::parameter sqlserver_command_service<DataBase>::create_parameter(
		implement_type& impl,string const& name, msado::data_type type,
		msado::parameter_direction direction, long size,long value)
	{
		return impl.create_parameter(name, type, direction, 
			static_cast<long>(size), _varinat_t(value));
	}

	template<typename DataBase>
	inline msado::parameter sqlserver_command_service<DataBase>::create_parameter(
		implement_type& impl,string const& name, msado::data_type type,
		msado::parameter_direction direction, long size,unsigned long value)
	{
		return impl.create_parameter(name, type, direction, 
			static_cast<long>(size), _variant_t(value));
	}

	template<typename DataBase>
	inline msado::parameter 
		sqlserver_command_service<DataBase>::create_parameter(
			implement_type& impl,string const& name, msado::data_type type,
			msado::parameter_direction direction, long size,double value)
	{
		return impl.create_parameter(name, type, direction, 
			static_cast<long>(size), _variant_t(value));
	}

	template<typename DataBase>
	inline msado::parameter sqlserver_command_service<DataBase>::create_parameter(
		implement_type& impl,string const& name, msado::data_type type,
		msado::parameter_direction direction, long size,string const& value)
	{
		return impl.create_parameter(name, type, direction,
			static_cast<long>(size), _variant_t(value));
	}

	template<typename DataBase>
	inline msado::parameter sqlserver_command_service<DataBase>::create_parameter(
		implement_type& impl,string const& name, msado::data_type type,
		msado::parameter_direction direction, long size,datetime const& value)
	{
		return impl.create_parameter(name, type, direction, 
			static_cast<long>(size), _variant_t(value));
	}

	template<typename DataBase>
	inline msado::parameters 
		sqlserver_command_service<DataBase>::get_parameters(
			implement_type const& impl) const
	{
		return impl.get_parameters();
	}
	
	template<typename DataBase>
	inline void sqlserver_command_service<DataBase>::set_command_type(
		implement_type& impl, msado::command_type plCmdType)
	{
		impl.set_command_type(plCmdType);
	}
	
	template<typename DataBase>
	inline msado::command_type 
		sqlserver_command_service<DataBase>::get_command_type(
			implement_type const& impl) const
	{
		return impl.get_command_type();
	}
	
	template<typename DataBase>
	inline msado::string sqlserver_command_service<DataBase>::get_name(
		implement_type const& impl) const
	{
		return impl.get_name();
	}
	
	template<typename DataBase>
	inline void sqlserver_command_service<DataBase>::set_name(
		implement_type& impl,string const& pbstrName)
	{
		impl.set_name(pbstrName);
	}
	
	template<typename DataBase>
	inline long sqlserver_command_service<DataBase>::get_state(
		implement_type const& impl) const
	{
		return impl.get_state();
	}
	
	template<typename DataBase>
	inline bool sqlserver_command_service<DataBase>::cancel(
		implement_type& impl)
	{
		return impl.cancel();
	}

	//void set_ref_command_stream(IUnknown * pvStream);
	//_variant_t get_command_stream() const;
	//void set_dialect(string const& pbstrDialect);
	//msado::string get_dialect() const;
	//void set_named_parameters(bool pfNamedParameters);
	//bool get_named_parameters() const;
}
