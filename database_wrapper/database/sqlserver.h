#pragma once

#include <atlbase.h>
#include <atlstr.h>
#include <guiddef.h>
#include <ATLComTime.h>

#include "sqlserver/detail/command.h"
#include "sqlserver/detail/recordset.h"
#include "sqlserver/detail/connection.h"

#include "detail/basic_service.h"
#include "detail/basic_object.h"
#include "detail/basic_command.h"
#include "detail/basic_recordset.h"
#include "detail/basic_connect.h"

#include "sqlserver/sqlserver_command_service.h"
#include "sqlserver/sqlserver_recordset_service.h"
#include "sqlserver/sqlserver_connect_service.h"

#include "sqlserver/sqlserver_connect.h"
#include "sqlserver/sqlserver_command.h"
#include "sqlserver/sqlserver_recordset.h"

namespace database
{
	using msado::cominit;

	class sqlserver
	{
	public:
		typedef sqlserver_connect<sqlserver> connect;
		typedef sqlserver_command<sqlserver> command;
		typedef sqlserver_recordset<sqlserver> recordset;

	public:
		explicit sqlserver(msado::string const& connstr = L"",
			msado::string const& user = L"",
			msado::string const& password = L"")
			: connstr_(connstr)
			, user_(user)
			, password_(password)
		{
		}

		sqlserver(sqlserver const& rhs)
			: connstr_(rhs.connstr_)
			, user_(rhs.user_)
			, password_(rhs.password_)
		{

		}
	public:
		msado::string get_connstr() const
		{
			return connstr_;
		}

		msado::string get_user() const
		{
			return user_;
		}

		msado::string get_password() const
		{
			return password_;
		}

		void set_connstr(msado::string const& connstr)
		{
			connstr_=connstr;
		}

		void set_user(msado::string const& user)
		{
			user_=user;
		}

		void set_password(msado::string const& password)
		{
			password_=password;
		}
	private:
		msado::string connstr_;
		msado::string user_;
		msado::string password_;
	};
}
