#pragma once

namespace database
{
	namespace detail
	{
		//
		//基本命令操作,sqlserver,sqlite等共有的操作。
		//
		template<typename DataBase, template<typename>class Service>
		class basic_recordset
			: public basic_object<Service<DataBase> >
		{
		public:
			basic_recordset()
				: basic_object<Service<DataBase> >()
			{
			}

			explicit basic_recordset(CLSID const& clsid)
				: basic_object<Service<DataBase> >(clsid)
			{

			}

			basic_recordset(DataBase const& db,msado::recordset& set)
				: basic_object<Service<DataBase> >(db,set.get())
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