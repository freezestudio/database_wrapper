#pragma once

namespace database
{
	namespace detail
	{
		//
		//基础类。表示 连接、命令、结果集等等对象。
		//
		template<typename BasicService>
		class basic_object
		{
		public:
			typedef BasicService service_type;
			typedef typename service_type::implement_type implement_type;
			typedef typename service_type::database_type database_type;
		protected:
			basic_object()
			{

			}

			//explicit basic_object(string const& filename)
			//	: implement_(filename)
			//{

			//}

			explicit basic_object(CLSID const& clsid)
				: implement_(clsid)
			{

			}

			explicit basic_object(database_type const& db,
				typename implement_type::interface_type* ptr)
				: service_(db)
				, implement_(ptr)
			{

			}

			basic_object(string const& connstr, CLSID const& clsid)
				: service_(connstr)
				, implement_(clsid)
			{

			}

			basic_object(string const& connstr,
				string const& user,
				string const& password,
				CLSID const& clsid)
				: service_(connstr, user, password)
				, implement_(clsid)
			{

			}

#ifdef __TEST__
		public:
			database_type const& get_database() const
			{
				return service_.get_database();
			}

			database_type& get_database()
			{
				return service_.get_database();
			}
#endif
		protected:
			service_type& get_service()
			{
				return service_;
			}

			service_type const& get_service() const
			{
				return service_;
			}

			implement_type& get_implement()
			{
				return implement_;
			}

			implement_type const& get_implement() const
			{
				return implement_;
			}
		protected:
			service_type service_;
			implement_type implement_;
		};
	}
}

namespace database
{
	///////////////////////////////////////////////////
	// object's class

	template<typename DataBase,
		template<typename>class SqlService = sqlserver_connect_service>
	class sqlserver_connect;
	template<typename DataBase,
		template<typename>class SqlService = sqlserver_command_service>
	class sqlserver_command;
	template<typename DataBase,
		template<typename>class SqlService = sqlserver_recordset_service>
	class sqlserver_recordset;

	template<typename DataBase,
		template<typename>class SQLiteService = sqlite_connect_service>
	class sqlite_connect;
	template<typename DataBase,
		template<typename>class SQLiteService = sqlite_command_service>
	class sqlite_command;
	template<typename DataBase,
		template<typename>class SQLiteService = sqlite_recordset_service>
	class sqlite_recordset;
}