#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的连接服务。
	//
	template<typename DataBase>
	class sqlserver_connect_service
		: public detail::basic_service<sqlserver_connect_service<DataBase> >
	{
	public:
		typedef msado::connection implement_type;
		typedef DataBase database_type;
	public:
		sqlserver_connect_service()
		{

		}

		explicit sqlserver_connect_service(database_type const& db)
			: database_(db)
		{

		}

		sqlserver_connect_service(sqlserver_connect_service const& rhs)
			: database_(rhs.database_)
		{

		}

		explicit sqlserver_connect_service(msado::string const& connstr)
			:database_(connstr)
		{

		}

		sqlserver_connect_service(msado::string const& connstr,
			msado::string const& user, msado::string const& password )
			:database_(connstr,user,password)
		{

		}

#ifdef __TEST__

	public:
		database_type const& get_database() const
		{
			return database_;
		}

		database_type& get_database()
		{
			return database_;
		}
#endif

	public:
		//bool open(implement_type& impl);

		//bool open(implement_type& impl, string const& connstr);
		//bool open(implement_type& impl, string const& connstr,
		//	string const& user, string const& password);

		bool open(implement_type& impl,
			string const& connstr=L"", 
			string const& user=L"",
			string const& password=L"",
			msado::connect_option options=msado::connect_unspecified);

		bool close(implement_type& impl);

		bool execute(implement_type& impl,msado::string const& cmdtext);		

		bool execute(implement_type& impl,msado::string const& cmdtext,
			long& result, msado::command_type options);

		bool execute(implement_type& impl, msado::string const& cmdtext,
			long& result, msado::command_type options,
			sqlserver_recordset<DataBase>& set);

		long begin_transaction(implement_type& impl);
		bool commit_transaction(implement_type& impl);
		bool rollback_transaction(implement_type& impl);
		void set_connection_timeout(implement_type& impl, long timeout = 15);

	private:
		bool opened(implement_type const& impl) const;
	private:
		database_type database_;
	};
}

namespace database
{
	//template<typename DataBase>
	//inline bool sqlserver_connect_service<DataBase>::open(implement_type& impl)
	//{
	//	return impl.open(database_.get_connstr(),
	//		database_.get_user(),database_.get_password());
	//}

	//template<typename DataBase>
	//inline bool sqlserver_connect_service<DataBase>::open(implement_type& impl,
	//	string const& connstr)
	//{
	//	database_.set_connstr(connstr);
	//	return impl.open(database_.get_connstr(),
	//		database_.get_user(), database_.get_password());
	//}

	//template<typename DataBase>
	//inline bool sqlserver_connect_service<DataBase>::open(implement_type& impl,
	//	string const& connstr, string const& user, string const& password)
	//{
	//	database_.set_connstr(connstr);
	//	database_.set_user(user);
	//	database_.set_password(password);

	//	return impl.open(database_.get_connstr(),
	//		database_.get_user(), database_.get_password());
	//}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::open(implement_type& impl,
		string const& connstr,string const& user, string const& password,
		msado::connect_option options)
	{
		if (!connstr.IsEmpty())database_.set_connstr(connstr);
		if(!user.IsEmpty())database_.set_user(user);
		if(!password.IsEmpty())database_.set_password(password);

		return impl.open(database_.get_connstr(),
			database_.get_user(), database_.get_password(),
			options);
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::close(implement_type& impl)
	{
		return impl.close();
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::execute(
		implement_type& impl, msado::string const& cmdtext)
	{
		long result = 0;
		impl.execute(cmdtext, result,msado::execute_no_records);
		return true;
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::execute(implement_type& impl,
		msado::string const& cmdtext,long& result,msado::command_type options)
	{
		msado::recordset set = impl.execute(cmdtext, result, options);
		return set.get()!=NULL;
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::execute(implement_type& impl,
		msado::string const& cmdtext,
		long& result, msado::command_type options,
		sqlserver_recordset<DataBase>& set)
	{
		msado::recordset vset = impl.execute(cmdtext, result, options);
		sqlserver_recordset<database::sqlserver> sql_set(database_, vset);
		set = sql_set;
		return true;
	}

	template<typename DataBase>
	inline long sqlserver_connect_service<DataBase>::begin_transaction(
		implement_type& impl)
	{
		return impl.begin_trans();
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::commit_transaction(
		implement_type& impl)
	{
		return impl.commit_trans();
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::rollback_transaction(
		implement_type& impl)
	{
		return impl.rollback_trans();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_connection_timeout(
		implement_type& impl, long timeout/* = 15*/)
	{
		impl.set_connection_timeout(static_cast<long>(timeout));
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::opened(
		implement_type const& impl) const
	{
		return impl.get_state() != msado::state_closed;
	}
}
