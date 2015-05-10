#include "stdafx.h"
#include <iostream>

#include "test_sqlserver.h"

sqlserver_test::sqlserver_test()
	: pconn(NULL)
	, pcmd(NULL)
	, pset(NULL)
{

}

void sqlserver_test::test_connect_create()
{
	if (pconn)
	{
		delete pconn;
		pconn = NULL;
	}

	pconn = new connect_test;

	ATLASSERT(pconn != NULL);
	ATLASSERT(pconn->get_database().get_connstr() == L"");

	delete pconn;
	pconn = NULL;

	database::sqlserver db(connection_string);
	msado::connection conn(msado::Clsid::ADODB_Connection());

	pconn = new connect_test(db, conn);

	ATLASSERT(pconn != NULL);
	ATLASSERT(db.get_connstr() == pconn->get_database().get_connstr());

	bool opened=pconn->open();
	ATLASSERT(!opened);

	opened = pconn->open(connection_string, L"", L"sql2014");
	ATLASSERT(opened);

	database::string connstr = conn.get_connection_string();
	if (opened)
	{
		pconn->close();
	}

	delete pconn;
	pconn = NULL;

	pconn = new connect_test(connection_string);

	ATLASSERT(pconn->get_database().get_connstr()==connection_string);

	delete pconn;
	pconn = NULL;
}

void sqlserver_test::test_connect_method()
{
	if (pconn)
	{
		delete pconn;
	}

	pconn = new connect_test(connection_string);
	pconn->get_database().set_password(L"sql2014");
	bool opened = pconn->open();
	if (opened)
	{
		long result = 0;
		database::sqlserver::recordset set;
		pconn->execute(L"select * from person", result, msado::cmd_text, set);

		database::string name = set.get_string_value(1);
		database::datetime dt = set.get_datetime_value(4);

	}
}

void sqlserver_test::test_connect_destroy()
{

}
