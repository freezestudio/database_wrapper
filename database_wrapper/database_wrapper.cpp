// sqlserver_sqlite_wrapper.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream>

#include "test_sqlserver.h"
#include "test_sqlite.h"

void TEST_sqlserver();
void TEST_sqlite();

// database : dbo.ado_test
// table    : dbo.person
// schema   : ID:int Name:varchar Uint:varchar Age:int Born:datetime Address:text
//
int _tmain(int argc, _TCHAR* argv[])
{
	database::cominit init;

	TEST_sqlserver();

	return 0;
}

void TEST_sqlserver()
{
	sqlserver_test test;

	//test.test_connect_create();
	//test.test_connect_method();
	//test.test_connect_destroy();

	//test.test_command_create();
	//test.test_command_method();
	//test.test_command_destroy();

	test.test_recordset_create();
	test.test_recordset_method();
	test.test_recordset_destroy();
}

void TEST_sqlite()
{

}

