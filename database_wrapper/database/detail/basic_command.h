#pragma once

namespace database
{
	namespace detail
	{
		//
		//基本命令操作,sqlserver,sqlite等共有的操作。
		//
		template<typename DataBase, template<typename>class Service>
		class basic_command
			: public basic_object<Service<DataBase> >
		{
		public:
			basic_command()
				: basic_object<Service<DataBase> >()
			{
			}

			basic_command(CLSID const& clsid)
				: basic_object<Service<DataBase> >(clsid)
			{

			}

			basic_command(DataBase const& db, msado::command& cmd)
				: basic_object<Service<DataBase> >(db,cmd.get())
			{

			}
		};
	}
}