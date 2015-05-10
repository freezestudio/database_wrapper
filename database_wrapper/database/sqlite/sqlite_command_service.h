#pragma once

#include "../detail/basic_service.h"
#include "detail/sqlite3_command.h"

namespace database
{
	template<typename DataBase>
	class sqlite_command_service
		: detail::basic_service<sqlite_command_service<DataBase> >
	{
	public:
		typedef sqlite_detail::sqlite3_command implement_type;
		typedef DataBase database_type;
	public:
		sqlite_command_service()
		{

		}
	private:
		database_type database_;
	};
}