#pragma once

namespace database
{
	namespace detail
	{
		//
		//基本连接操作,sqlserver,sqlite等共有的操作。
		//
		template<typename DataBase, template<typename>class Service>
		class basic_connect 
			: public basic_object<Service<DataBase> >
		{
		public:
			basic_connect()
				: basic_object<Service<DataBase> >()
			{
			}

			basic_connect(DataBase const& db, msado::connection const& conn)
				: basic_object<Service<DataBase> >(db,conn.get())
			{

			}

			explicit basic_connect(CLSID const& clsid)
				: basic_object<Service<DataBase> >(clsid)
			{

			}

			basic_connect(string const& connstr, CLSID const& clsid)
				: basic_object<Service<DataBase> >(connstr,clsid)
			{

			}

			basic_connect(string const& connstr, 
				string const& user, 
				string const& password, 
				CLSID const& clsid)
				: basic_object<Service<DataBase> >(connstr, user,password,clsid)
			{

			}
		public:
			bool close()
			{
				return get_service().close(get_implement());
			}
		};
	}
}
