#pragma once
#include "sqlite/detail/sqlite3_command.h"
#include "sqlite/detail/sqlite3_recordset.h"
#include "sqlite/detail/sqlite3_connection.h"

#include "detail/basic_service.h"
#include "detail/basic_object.h"
#include "detail/basic_command.h"
#include "detail/basic_recordset.h"
#include "detail/basic_connect.h"

#include "sqlite/sqlite_command_service.h"
#include "sqlite/sqlite_recordset_service.h"
#include "sqlite/sqlite_connect_service.h"

#include "sqlite/sqlite_connect.h"
#include "sqlite/sqlite_command.h"
#include "sqlite/sqlite_recordset.h"

namespace database
{
	class sqlite
	{
	public:
		typedef sqlite_connect<sqlite> connect;
	public:
		sqlite()
		{

		}
	private:
		detail::string db_name_;
	};
}
