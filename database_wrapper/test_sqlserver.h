#pragma once

#define __TEST__

#include "database/sqlserver.h"

const static wchar_t* const connection_string = 
	L"Provider=SQLOLEDB.1;"
	L"Persist Security Info=True;"
	L"User ID=sa;"
	L"Initial Catalog=ado_test;"
	L"Data Source=127.0.0.1";

class sqlserver_test
{
public:
	class connect_test;
	class command_test;
	class recordset_test;
public:
	sqlserver_test();
public:
	void test_connect_create();
	void test_connect_method();
	void test_connect_destroy();

	void test_command_create();
	void test_command_method();
	void test_command_destroy();

	void test_recordset_create();
	void test_recordset_method();
	void test_recordset_destroy();
private:
	connect_test* pconn;
	command_test* pcmd;
	recordset_test* pset;
};

typedef database::sqlserver::connect adoconn;

class sqlserver_test::connect_test : public adoconn
{
public:
	explicit connect_test(CLSID const& clsid = msado::Clsid::ADODB_Connection())
		: adoconn(clsid)
	{
	}

	connect_test(database::sqlserver const& db, msado::connection& conn)
		: adoconn(db, conn)
	{

	}

	explicit connect_test(msado::string const& connstr,
		CLSID const& clsid = msado::Clsid::ADODB_Connection())
		: adoconn(connstr, clsid)
	{
	}

	connect_test(msado::string const& connstr,
		msado::string const& user,
		msado::string const& password,
		CLSID const& clsid = msado::Clsid::ADODB_Connection())
		: adoconn(connstr, user, password, clsid)
	{
	}
};

typedef database::sqlserver::command adocmd;

class sqlserver_test::command_test : public adocmd
{
public:
	explicit command_test(
		CLSID const& clsid = msado::Clsid::ADODB_Command())
		: adocmd(clsid)
	{
	}

	command_test(database::sqlserver const& db, msado::command& cmd)
		: adocmd(db,cmd)
	{

	}
};

typedef database::sqlserver::recordset adoset;

class sqlserver_test::recordset_test : public adoset
{
public:
	explicit recordset_test(
		CLSID const& clsid = msado::Clsid::ADODB_Recordset())
		: adoset(clsid)
	{
	}

	recordset_test(database::sqlserver const& db,msado::recordset& set)
		: adoset(db,set)
	{

	}
};