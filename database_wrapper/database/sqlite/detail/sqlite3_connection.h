#pragma once

#include "sqlite3.h"

namespace sqlite_detail
{
	class sqlite3_connection
	{
	public:
		sqlite3_connection()
		{
			sqlite3_open("", &sqlite_);
		}
		sqlite3_connection(sqlite3_connection const&)
		{

		}
	private:
		sqlite3* sqlite_;
	};
}