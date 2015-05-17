#include "stdafx.h"
#include <iostream>

#include "test_sqlserver.h"

sqlserver_test::sqlserver_test()
	: pconn(NULL)
	, pcmd(NULL)
	, pset(NULL)
{

}

//////////////////////////////////////////////////////////
// conection

void sqlserver_test::test_connect_create()
{
	if (pconn)
	{
		pconn->close();
		delete pconn;
		pconn = NULL;
	}

	pconn = new connect_test;

	ATLASSERT(pconn != NULL);
	ATLASSERT(pconn->get_database().get_connstr() == L"");

	delete pconn;
	pconn = NULL;

	//////////////////////////////////////////////////////////

	database::sqlserver db(connection_string);
	msado::connection conn(msado::Clsid::ADODB_Connection());

	pconn = new connect_test(db, conn);

	ATLASSERT(pconn != NULL);
	ATLASSERT(db.get_connstr() == pconn->get_database().get_connstr());

	bool opened = pconn->open();
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

	/////////////////////////////////////////////////////////

	pconn = new connect_test(connection_string);

	ATLASSERT(pconn->get_database().get_connstr() == connection_string);
	database::sqlserver& sqlserver = pconn->get_database();

	sqlserver.set_password(L"sql2014");
	ATLASSERT(pconn->get_database().get_password() == sqlserver.get_password());

	pconn->open();

	ATLASSERT(pconn->get_internal_ptr());
	long result = 0;
	database::sqlserver_recordset<database::sqlserver> set =
		pconn->execute(L"select * from person", result, msado::cmd_text);
	database::string name = set.get_string_value(L"Name");

	set.close();
	pconn->close();

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
		database::sqlserver::recordset set =
			pconn->execute(L"select * from person", result, msado::cmd_text);

		database::string name = set.get_string_value(1);
		database::datetime dt = set.get_datetime_value(4);

		bool badd_new = set.supports(msado::cursor_option::add_new);
		bool bmove = set.supports(msado::cursor_option::move_pervious);
		bool bupdate = set.supports(msado::cursor_option::update);
		bool bupdatebatch = set.supports(msado::cursor_option::update_batch);

		msado::edit_mode medit = set.get_edit_mode();
		bool can_add_new = set.add_new();
		medit = set.get_edit_mode();

		//VARIANT var;
		//VariantInit(&var);
		//set.get_internal_ptr()->GetRows(-1, vtMissing, vtMissing, &var);

		msado::edit_add;//2
		msado::edit_none;//0
	}
}

void sqlserver_test::test_connect_destroy()
{
	if (pconn)
	{
		pconn->close();
		delete pconn;
		pconn = NULL;
	}
}

//////////////////////////////////////////////////////////
// command

void sqlserver_test::test_command_create()
{

}

void sqlserver_test::test_command_method()
{

}
void sqlserver_test::test_command_destroy()
{

}

//////////////////////////////////////////////////////////
// recordset

void sqlserver_test::test_recordset_create()
{
	if (pset)
	{
		pset->close();
		pset = NULL;
	}

	pconn = new connect_test(connection_string, L"", L"sql2014");
	pconn->open();

	msado::cursor_location cl = pconn->get_cursor_location();

	pset = new recordset_test;
	bool set_opened = pset->open(L"person", *pconn);

	//Modify
	//pset->set_value(1, L"张三");
	//pset->update();

	//AddNew
	bool can_add = pset->supports(msado::cursor_option::add_new);
	bool b_add_new = pset->add_new();
	if (pset->get_edit_mode() == msado::edit_add)
	{
		pset->set_value(L"Name", L"王五");
		pset->set_value(L"Unit", L"保卫部");
		pset->set_value(L"Age", long(30));
		pset->set_value(4, L"1990-3-11");
		pset->set_value(L"Address", L"宁夏银川");
		try
		{
			bool bupdate=pset->update();
		}
		catch (_com_error& e)
		{
			_bstr_t dest=e.Description();
			_bstr_t msg = e.ErrorMessage();
		}
		//bool batch=pset->update_batch(msado::affect::affect_all);
	}

	pset->close();
	pconn->close();

	delete pset;
	delete pconn;
	pset = NULL;
	pconn = NULL;

}

void sqlserver_test::test_recordset_method()
{

}

void sqlserver_test::test_recordset_destroy()
{

}
