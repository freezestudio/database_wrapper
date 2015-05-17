#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的连接操作
	//
	template<typename DataBase,
		template<typename>class SqlService/* = sqlserver_connect_service*/>
	class sqlserver_connect : public detail::basic_connect<DataBase, SqlService>
	{
	public:
		explicit sqlserver_connect(
			CLSID const& clsid = msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(clsid)
		{
		}

		sqlserver_connect(DataBase const& db, msado::connection& conn)
			: detail::basic_connect<DataBase, SqlService>(db, conn)
		{

		}

		explicit sqlserver_connect(msado::string const& connstr,
			CLSID const& clsid = msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(connstr, clsid)
		{
		}

		sqlserver_connect(msado::string const& connstr,
			msado::string const& user,
			msado::string const& password,
			CLSID const& clsid = msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(
				connstr, user, password, clsid)
		{
		}

	public:
		//bool open();
		//bool open(string const& connstr);
		//bool open(string const& connstr,
		//	string const& user, string const& password);

		bool open(string const& connstr = L"",
			string const& user = L"",
			string const& password = L"",
			msado::connect_option options = msado::connect_unspecified);

		bool execute(msado::string const& cmdtext);
		//bool execute(msado::string const& cmdtext,
		//	long& result, msado::command_type options);

		sqlserver_recordset<DataBase> execute(msado::string const& cmdtext,
			long& result, msado::command_type options);

		long begin_transaction();
		bool commit_transaction();
		bool rollback_transaction();

		database::string get_connection_string() const;
		void set_connection_string(string const& pbstr);

		long get_command_timeout() const;
		void set_command_timeout(long plTimeout = 30 /*seconds*/);

		long get_connection_timeout() const;
		void set_connection_timeout(long timeout = 15 /*seconds*/);

		database::string get_version() const;

		//msado::errors get_errors() const;

		database::string get_default_database() const;
		void set_default_database(string const& pbstr);

		msado::isolation_level get_isolation_level() const;
		void set_isolation_level(msado::isolation_level Level);

		long get_attributes() const;
		void set_attributes(long plAttr);

		msado::cursor_location get_cursor_location() const;
		void set_cursor_location(msado::cursor_location plCursorLoc);

		msado::connect_mode get_mode() const;
		void set_mode(msado::connect_mode plMode);
		database::string get_provider() const;
		void set_provider(database::string const& pbstr);
		long get_state() const;
		sqlserver_recordset<DataBase> open_schema(msado::schema Schema);

		bool cancel();
	public:
		typename implement_type::interface_type const* get_internal_ptr() const;
		typename implement_type::interface_type * get_internal_ptr();
	};
}

namespace database
{
	//template<typename DataBase, template<typename> class SqlService>
	//inline bool sqlserver_connect<DataBase, SqlService>::open()
	//{
	//	return detail::basic_connect<DataBase, SqlService>::open();
	//}

	//template<typename DataBase,template<typename> class SqlService>
	//inline bool sqlserver_connect<DataBase,SqlService>::open(string const& connstr)
	//{
	//	return get_service().open(get_implement(),connstr);
	//}

	//template<typename DataBase, template<typename> class SqlService>
	//inline bool sqlserver_connect<DataBase, SqlService>::open(string const& connstr,
	//	string const& user, string const& password)
	//{
	//	return get_service().open(get_implement(),
	//		connstr, user, password);
	//}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::open(string const& connstr,
		string const& user, string const& password, msado::connect_option options)
	{
		return get_service().open(get_implement(),
			connstr, user, password, options);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::execute(
		msado::string const& cmdtext)
	{
		return get_service().execute(get_implement(), cmdtext);
	}

	//template<typename DataBase, template<typename> class SqlService>
	//inline bool sqlserver_connect<DataBase, SqlService>::execute(
	//	msado::string const& cmdtext, long& result, msado::command_type options)
	//{
	//	return get_service().execute(get_implement(),
	//		cmdtext, result, options);
	//}

	template<typename DataBase, template<typename> class SqlService>
	inline sqlserver_recordset<DataBase>
		sqlserver_connect<DataBase, SqlService>::execute(
			msado::string const& cmdtext,
			long& result, msado::command_type options)
	{
		return get_service().execute(get_implement(),
			cmdtext, result, options);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline long sqlserver_connect<DataBase, SqlService>::begin_transaction()
	{
		return get_service().begin_transaction(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::commit_transaction()
	{
		return get_service().commit_transaction(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::rollback_transaction()
	{
		return get_service().roolback_transaction(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline database::string
		sqlserver_connect<DataBase, SqlService>::get_connection_string() const
	{
		return get_service().get_connection_string(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_connection_string(
		string const& pbstr)
	{
		get_service().set_connection_string(get_implement(), pbstr);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline long sqlserver_connect<DataBase, SqlService>::get_command_timeout() const
	{
		return get_service().get_command_timeout(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_command_timeout(
		long plTimeout/* = 30 seconds*/)
	{
		return get_service().set_command_timeout(get_implement(), plTimeout);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline long
		sqlserver_connect<DataBase, SqlService>::get_connection_timeout() const
	{
		return get_service().get_connection_timeout(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_connection_timeout(
		long timeout/* = 15 seconds*/)
	{
		return get_service().set_connection_timeout(get_implement(), timeout);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline database::string
		sqlserver_connect<DataBase, SqlService>::get_version() const
	{
		return get_service().get_version(get_implement());
	}

	//template<typename DataBase, template<typename> class SqlService>
	//inline msado::errors sqlserver_connect<DataBase, SqlService>::get_errors() const
	//{
	//
	//}

	template<typename DataBase, template<typename> class SqlService>
	inline database::string
		sqlserver_connect<DataBase, SqlService>::get_default_database() const
	{
		return get_service().get_default_database(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_default_database(
		string const& pbstr)
	{
		return get_service().set_default_database(get_implement(), pbstr);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline msado::isolation_level
		sqlserver_connect<DataBase, SqlService>::get_isolation_level() const
	{
		return get_service().get_isolation_level(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_isolation_level(
		msado::isolation_level Level)
	{
		return get_service().set_isolation_level(get_implement(), Level);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline long sqlserver_connect<DataBase, SqlService>::get_attributes() const
	{
		return get_service().get_attributes(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_attributes(
		long plAttr)
	{
		return get_service().set_attributes(get_implement(), plAttr);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline msado::cursor_location
		sqlserver_connect<DataBase, SqlService>::get_cursor_location() const
	{
		return get_service().get_cursor_location(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_cursor_location(
		msado::cursor_location plCursorLoc)
	{
		return get_service().set_cursor_location(get_implement(), plCursorLoc);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline msado::connect_mode
		sqlserver_connect<DataBase, SqlService>::get_mode() const
	{
		return get_service().get_mode(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_mode(
		msado::connect_mode plMode)
	{
		return get_service().set_mode(get_implement(), plMode);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline database::string
		sqlserver_connect<DataBase, SqlService>::get_provider() const
	{
		return get_service().get_provider(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline void sqlserver_connect<DataBase, SqlService>::set_provider(
		database::string const& pbstr)
	{
		return get_service().set_provider(get_implement(), pbstr);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline long sqlserver_connect<DataBase, SqlService>::get_state() const
	{
		return get_service().get_state(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline sqlserver_recordset<DataBase>
		sqlserver_connect<DataBase, SqlService>::open_schema(msado::schema Schema)
	{
		return get_service().open_schema(get_implement(), Schema);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::cancel()
	{
		return get_service().cancel(get_implement());
	}

	template<typename DataBase, template<typename> class SqlService>
	inline typename sqlserver_connect<DataBase, SqlService>::
		implement_type::interface_type const*
		sqlserver_connect<DataBase, SqlService>::get_internal_ptr() const
	{
		return get_implement().get();
	}

	template<typename DataBase, template<typename> class SqlService>
	inline typename sqlserver_connect<DataBase, SqlService>::
		implement_type::interface_type*
		sqlserver_connect<DataBase, SqlService>::get_internal_ptr()
	{
		return get_implement().get();
	}
}