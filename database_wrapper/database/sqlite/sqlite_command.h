#pragma once

#include "../detail/basic_command.h"
#include "sqlite_command_service.h"

namespace database
{
	template<typename DataBase,
		template<typename>class SQLiteService/* = sqlite_command_service*/>
	class sqlite_command
		: public detail::basic_command<DataBase, SQLiteService>
	{

	};
}