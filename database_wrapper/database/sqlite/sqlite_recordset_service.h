#pragma once

#include "detail/sqlite3_recordset.h"
#include "../detail/basic_service.h"

namespace database
{
	template<typename DataBase>
	class sqlite_recordset_service
		: public detail::basic_service<sqlite_recordset_service<DataBase> >
	{
	public:
		typedef sqlite_detail::sqlite3_recordset implement_type;
		typedef DataBase database_type;
	public:
		sqlite_recordset_service()
		{

		}
	private:
		database_type database_;
	};
}