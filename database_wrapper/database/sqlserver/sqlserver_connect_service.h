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
			msado::string const& user, msado::string const& password)
			:database_(connstr, user, password)
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
			string const& connstr = L"",
			string const& user = L"",
			string const& password = L"",
			msado::connect_option options = msado::connect_unspecified);

		bool close(implement_type& impl);

		bool execute(implement_type& impl, msado::string const& cmdtext);

		//bool execute(implement_type& impl, msado::string const& cmdtext,
		//	long& result, msado::command_type options);

		sqlserver_recordset<DataBase> execute(implement_type& impl,
			msado::string const& cmdtext,
			long& result, msado::command_type options);

		long begin_transaction(implement_type& impl);
		bool commit_transaction(implement_type& impl);
		bool rollback_transaction(implement_type& impl);

		database::string get_connection_string(implement_type const& impl) const;
		void set_connection_string(implement_type& impl, string const& pbstr);

		long get_command_timeout(implement_type const& impl) const;
		void set_command_timeout(implement_type& impl, long plTimeout = 30/*seconds*/);

		long get_connection_timeout(implement_type const& impl) const;
		void set_connection_timeout(implement_type& impl, long timeout = 15);

		database::string get_version(implement_type const& impl) const;

		//msado::errors get_errors(implement_type const& impl) const;

		database::string get_default_database(implement_type const& impl) const;
		void set_default_database(implement_type& impl, string const& pbstr);

		msado::isolation_level get_isolation_level(implement_type const& impl) const;
		void set_isolation_level(implement_type& impl, msado::isolation_level Level);

		long get_attributes(implement_type const& impl) const;
		void set_attributes(implement_type& impl, long plAttr);

		msado::cursor_location get_cursor_location(implement_type const& impl) const;
		void set_cursor_location(implement_type& impl, msado::cursor_location plCursorLoc);

		msado::connect_mode get_mode(implement_type const& impl) const;
		void set_mode(implement_type& impl, msado::connect_mode plMode);
		database::string get_provider(implement_type const& impl) const;
		void set_provider(implement_type& impl, database::string const& pbstr);
		long get_state(implement_type const& impl) const;
		sqlserver_recordset<DataBase> open_schema(
			implement_type& impl, msado::schema Schema);

		bool cancel(implement_type& impl);

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
		string const& connstr, string const& user, string const& password,
		msado::connect_option options)
	{
		if (!connstr.IsEmpty())database_.set_connstr(connstr);
		if (!user.IsEmpty())database_.set_user(user);
		if (!password.IsEmpty())database_.set_password(password);

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
		impl.execute(cmdtext, result, msado::execute_no_records);
		return true;
	}

	//template<typename DataBase>
	//inline bool sqlserver_connect_service<DataBase>::execute(implement_type& impl,
	//	msado::string const& cmdtext, long& result, msado::command_type options)
	//{
	//	msado::recordset set = impl.execute(cmdtext, result, options);
	//	return set.get() != NULL;
	//}

	template<typename DataBase>
	inline sqlserver_recordset<DataBase>
		sqlserver_connect_service<DataBase>::execute(implement_type& impl,
			msado::string const& cmdtext,
			long& result, msado::command_type options)
	{
		msado::recordset vset = impl.execute(cmdtext, result, options);
		sqlserver_recordset<database::sqlserver> sql_set(database_, vset);
		return sql_set;
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
	inline database::string
		sqlserver_connect_service<DataBase>::get_connection_string(
			implement_type const& impl) const
	{
		return impl.get_connection_string();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_connection_string(
		implement_type& impl, string const& pbstr)
	{
		impl.set_connection_string(pbstr);
	}

	template<typename DataBase>
	inline long sqlserver_connect_service<DataBase>::get_command_timeout(
		implement_type const& impl) const
	{
		return impl.get_command_timeout();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_command_timeout(
		implement_type& impl, long plTimeout/* = 30 seconds*/)
	{
		impl.set_command_timeout(plTimeout);
	}

	template<typename DataBase>
	inline long sqlserver_connect_service<DataBase>::get_connection_timeout(
		implement_type const& impl) const
	{
		return impl.get_connection_timeout();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_connection_timeout(
		implement_type& impl, long timeout/* = 15 seconds*/)
	{
		impl.set_connection_timeout(static_cast<long>(timeout));
	}

	template<typename DataBase>
	inline database::string sqlserver_connect_service<DataBase>::get_version(
		implement_type const& impl) const
	{
		return impl.get_version();
	}

	//template<typename DataBase>
	//inline msado::errors get_errors(implement_type const& impl) const
	//{
	//
	//}

	template<typename DataBase>
	inline database::string
		sqlserver_connect_service<DataBase>::get_default_database(
			implement_type const& impl) const
	{
		return impl.get_default_database();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_default_database(
		implement_type& impl, string const& pbstr)
	{
		impl.set_default_database(pbstr);
	}

	template<typename DataBase>
	inline msado::isolation_level
		sqlserver_connect_service<DataBase>::get_isolation_level(
			implement_type const& impl) const
	{
		return impl.get_isolation_level();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_isolation_level(
		implement_type& impl, msado::isolation_level Level)
	{
		impl.set_isolation_level(Level);
	}

	template<typename DataBase>
	inline long sqlserver_connect_service<DataBase>::get_attributes(
		implement_type const& impl) const
	{
		return impl.get_attributes();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_attributes(
		implement_type& impl, long plAttr)
	{
		impl.set_attributes(plAttr);
	}

	template<typename DataBase>
	inline msado::cursor_location
		sqlserver_connect_service<DataBase>::get_cursor_location(
			implement_type const& impl) const
	{
		return impl.get_cursor_location();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_cursor_location(
		implement_type& impl, msado::cursor_location plCursorLoc)
	{
		impl.set_cursor_location(plCursorLoc);
	}

	template<typename DataBase>
	inline msado::connect_mode sqlserver_connect_service<DataBase>::get_mode(
		implement_type const& impl) const
	{
		return impl.get_mode();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_mode(
		implement_type& impl, msado::connect_mode plMode)
	{
		impl.set_mode(plMode);
	}

	template<typename DataBase>
	inline database::string sqlserver_connect_service<DataBase>::get_provider(
		implement_type const& impl) const
	{
		return impl.get_provider();
	}

	template<typename DataBase>
	inline void sqlserver_connect_service<DataBase>::set_provider(
		implement_type& impl, database::string const& pbstr)
	{
		impl.set_provider(pbstr);
	}

	template<typename DataBase>
	inline long sqlserver_connect_service<DataBase>::get_state(
		implement_type const& impl) const
	{
		return impl.get_state();
	}

	template<typename DataBase>
	inline sqlserver_recordset<DataBase>
		sqlserver_connect_service<DataBase>::open_schema(
			implement_type& impl, msado::schema Schema)
	{
		msado::recordset set = impl.open_schema(Schema);
		return sqlserver_recordset<DataBase>(database_, set);
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::cancel(implement_type& impl)
	{
		return impl.cancel();
	}

	template<typename DataBase>
	inline bool sqlserver_connect_service<DataBase>::opened(
		implement_type const& impl) const
	{
		return impl.get_state() != msado::state_closed;
	}
}
