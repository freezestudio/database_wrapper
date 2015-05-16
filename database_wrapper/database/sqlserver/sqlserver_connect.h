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
			CLSID const& clsid=msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(clsid)
		{
		}

		sqlserver_connect(DataBase const& db,msado::connection const& conn)
			: detail::basic_connect<DataBase,SqlService>(db,conn)
		{

		}

		explicit sqlserver_connect(msado::string const& connstr,
			CLSID const& clsid = msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(connstr,clsid)
		{
		}

		sqlserver_connect(msado::string const& connstr,
			msado::string const& user,
			msado::string const& password,
			CLSID const& clsid = msado::Clsid::ADODB_Connection())
			: detail::basic_connect<DataBase, SqlService>(connstr,user,password,clsid)
		{
		}

	public:
		//bool open();
		//bool open(string const& connstr);
		//bool open(string const& connstr, string const& user, string const& password);
		bool open(string const& connstr=L"",
			string const& user=L"",
			string const& password=L"",
			msado::connect_option options=msado::connect_unspecified);

		bool execute(msado::string const& cmdtext);
		bool execute(msado::string const& cmdtext,
			long& result, msado::command_type options);

		bool execute(msado::string const& cmdtext,
			long& result, msado::command_type options,
			sqlserver_recordset<DataBase>& set);

		long begin_transaction();
		bool commit_transaction();
		bool rollback_transaction();
		void set_connection_timeout(long timeout = 15);

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
			connstr,user,password,options);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::execute(
			msado::string const& cmdtext)
	{
		return get_service().execute(get_implement(), cmdtext);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::execute(
			msado::string const& cmdtext, long& result, msado::command_type options)
	{
		return get_service().execute(get_implement(),
			cmdtext, result, options);
	}

	template<typename DataBase, template<typename> class SqlService>
	inline bool sqlserver_connect<DataBase, SqlService>::execute(
		msado::string const& cmdtext,
		long& result, msado::command_type options,
		sqlserver_recordset<DataBase>& set)
	{
		return get_service().execute(get_implement(),
			cmdtext, result, options, set);
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
	inline void sqlserver_connect<DataBase, SqlService>::set_connection_timeout(
		long timeout/* = 15*/)
	{
		return get_service().set_connection_timeout(get_implement(), timeout);
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