#pragma once

#include "../detail/basic_connect.h"
#include "sqlite_connect_service.h"

namespace database
{
	//
	//表示用于SQLITE的连接操作
	//
	template<typename DataBase,
		template<typename>class SQLiteService/* = sqlite_connect_service*/>
	class sqlite_connect
		: public detail::basic_connect<DataBase, SQLiteService>
	{

	};
}
