#pragma once

#include "../detail/basic_recordset.h"
#include "sqlite_recordset_service.h"

namespace database
{
	template<typename DataBase,
		template<typename>class SQLiteService/* = sqlite_recordset_service*/>
	class sqlite_recordset
		: public detail::basic_recordset<DataBase,SQLiteService>
	{

	};
}