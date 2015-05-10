#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的命令操作
	//
	template<typename DataBase, 
		template<typename>class SqlService/* = sqlserver_command_service*/>
	class sqlserver_command
		: public detail::basic_command<DataBase, SqlService>
	{
	public:
		explicit sqlserver_command(
			CLSID const& clsid=msado::Clsid::ADODB_Command())
			: detail::basic_command<DataBase, SqlService>(clsid)
		{
		}

		sqlserver_command(DataBase const& db, msado::command const& cmd)
			: detail::basic_command<DataBase,SqlService>(db,cmd)
		{

		}
	public:
		typename implement_type::interface_type const* get_internal_ptr() const;
		typename implement_type::interface_type * get_internal_ptr();
	};
}

namespace database
{

}