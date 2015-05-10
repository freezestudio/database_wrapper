#pragma once

namespace database
{
	namespace detail
	{
		typedef ATL::CString string;
		typedef ATL::COleDateTime datetime;

		//
		//对象的服务，实现对象的方法。
		//
		template<typename Service>
		class basic_service
		{
		public:
			basic_service()
			{
			}
		protected:
			bool is_opened()
			{
				Service* service = static_cast<Service*>(this);
				return service->opened();
			}
		};

	}
}

namespace database
{
	/////////////////////////////////////////////////
	// types

	using detail::string;
	using detail::datetime;

	////////////////////////////////////////////////
	// Service's class

	template<typename DataBase>	class sqlserver_connect_service;
	template<typename DataBase>	class sqlserver_command_service;
	template<typename DataBase>	class sqlserver_recordset_service;

	template<typename DataBase>	class sqlite_connect_service;
	template<typename DataBase>	class sqlite_command_service;
	template<typename DataBase>	class sqlite_recordset_service;
}
