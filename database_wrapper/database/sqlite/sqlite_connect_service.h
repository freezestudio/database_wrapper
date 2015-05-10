#pragma once

#include "../detail/basic_service.h"
#include "detail/sqlite3_connection.h"

namespace database
{
	//
	//表示用于SQLITE的连接服务。
	//
	template<typename DataBase>
	class sqlite_connect_service
		: detail::basic_service<sqlite_connect_service<DataBase> >
	{
	public:
		typedef sqlite_detail::sqlite3_connection implement_type;
		typedef DataBase database_type;
	public:
		sqlite_connect_service()
			: database_(DataBase::instance())
		{

		}
	public:
		bool open(implement_type& impl);
		bool close(implement_type& impl);
	private:
		database_type database_;
	};
}
