// sqlserver_sqlite_wrapper.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream>

#include "test_sqlserver.h"
#include "test_sqlite.h"

void TEST_sqlserver();
void TEST_sqlite();

int _tmain(int argc, _TCHAR* argv[])
{
	database::cominit init;

	msado::connection conn(msado::Clsid::ADODB_Connection());
	conn.open(connection_string, L"", L"sql2014");
	msado::command cmd(msado::Clsid::ADODB_Command());
	cmd.set_active_connection(_variant_t(conn.get()));
	cmd.set_command_text(L"select * from person");
	_variant_t cmdstream=cmd.get_command_stream();

	TEST_sqlserver();

	return 0;
}

void TEST_sqlserver()
{
	sqlserver_test test;
	test.test_connect_create();
	test.test_connect_method();
}

void TEST_sqlite()
{

}

