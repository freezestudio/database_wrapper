#include "recordset.h"

/////////////////////////////////////////////////
// ctor dtor internal method

msado::recordset::recordset()
	: cimpl<recordset,_RecordsetPtr>()
{
}

msado::recordset::recordset(recordset const& rhs)
	: cimpl<recordset,_RecordsetPtr>(rhs)
{
}

msado::recordset::recordset(recordset::interface_type* ptr)
	: cimpl<recordset, _RecordsetPtr>(ptr)
{
}

msado::recordset::recordset(CLSID const& clsid)
	: cimpl<recordset, _RecordsetPtr>(clsid)
{
}

msado::recordset::~recordset()
{
}

bool msado::recordset::enable_event(recordset_event* pevent, bool enable/* = true*/)
{
	//一旦启动事件，将产生了一错误中断，还没查明原因！！！
	return pevent->enable_event(enable);
}

//////////////////////////////////////////////////////////////
// public method

msado::properties msado::recordset::get_properties() const
{
	properties prop;
	ptr_->get_Properties(prop.get_address_of());
	return prop;
}

msado::position msado::recordset::get_absolute_position() const
{
	PositionEnum pl;
	ptr_->get_AbsolutePosition(&pl);
	return static_cast<position>(pl);
}

void msado::recordset::set_absolute_position(position pl)
{
	ptr_->put_AbsolutePosition(static_cast<PositionEnum>(pl));
}

//TODO
void msado::recordset::set_ref_active_connection(msado::connection& pvar)
{
	ptr_->putref_ActiveConnection(
		reinterpret_cast<IDispatch*>(pvar.get()));
}

void msado::recordset::set_active_connection(_variant_t const& pvar)
{
	ptr_->put_ActiveConnection(pvar);
}

//TODO
msado::connection msado::recordset::get_ref_active_connection() const
{	
	VARIANT var;
	VariantInit(&var);
	ptr_->get_ActiveConnection(&var);
	if (var.vt==VT_DISPATCH)
	{
		connection conn(reinterpret_cast<connection::interface_type*>(var.pdispVal));
		return conn;
	}
	return connection();
}

msado::string msado::recordset::get_active_connection() const
{
	string strconn;

	VARIANT var;
	VariantInit(&var);

	ptr_->get_ActiveConnection(&var);
	if (var.vt == VT_BSTR)
	{
		strconn = var.bstrVal;
	}
	return strconn;
}

bool msado::recordset::get_BOF() const
{
	VARIANT_BOOL vb=0;
	ptr_->get_BOF(&vb);
	return vb == VARIANT_TRUE;
}

_variant_t msado::recordset::get_bookmark() const
{
	VARIANT bookmark;
	VariantInit(&bookmark);

	ptr_->get_Bookmark(&bookmark);

	return _variant_t(bookmark, false);
}

void msado::recordset::set_bookmark(_variant_t const& pvBookmark)
{
	ptr_->put_Bookmark(pvBookmark);
}

long msado::recordset::get_cache_size() const
{
	long result = 0;
	ptr_->get_CacheSize(&result);
	return result;
}

void msado::recordset::set_cache_size(long pl)
{
	ptr_->put_CacheSize(pl);
}

msado::cursor_type msado::recordset::get_cursor_type() const
{
	CursorTypeEnum ct;
	ptr_->get_CursorType(&ct);
	return static_cast<cursor_type>(ct);
}

void msado::recordset::set_cursor_type(cursor_type plCursorType/* = open_dynamic*/)
{
	ptr_->put_CursorType(static_cast<CursorTypeEnum>(plCursorType));
}

bool msado::recordset::get_EOF() const
{
	VARIANT_BOOL vb = 0;
	ptr_->get_EOF(&vb);
	return vb == VARIANT_TRUE;
}

msado::fields msado::recordset::get_fields() const
{
	fields fs;
	ptr_->get_Fields(fs.get_address_of());
	return fs;
}

msado::fields msado::recordset::get_fields()
{
	fields fs;
	ptr_->get_Fields(fs.get_address_of());
	return fs;
}

msado::lock_type msado::recordset::get_lock_type() const
{
	LockTypeEnum lt;
	ptr_->get_LockType(&lt);
	return static_cast<lock_type>(lt);
}

void msado::recordset::set_lock_type(lock_type plLockType/* = lock_read_only*/)
{
	ptr_->put_LockType(static_cast<LockTypeEnum>(plLockType));
}

long msado::recordset::get_max_records() const
{
	long result = 0;
	ptr_->get_MaxRecords(&result);
	return result;
}

void msado::recordset::set_max_records(long plMaxRecords)
{
	ptr_->put_MaxRecords(plMaxRecords);
}

long msado::recordset::get_record_count() const
{
	long result = 0;
	ptr_->get_RecordCount(&result);
	return result;
}

void msado::recordset::set_ref_source(command const& pvSource)
{
	ptr_->putref_Source(reinterpret_cast<IDispatch*>(pvSource.get()));
}

void msado::recordset::set_source(string const& pvSource)
{
	_bstr_t bstr = detail::string_to_bstr_t(pvSource);
	ptr_->put_Source(bstr.GetBSTR());
}

//TODO
msado::command msado::recordset::get_ref_source() const
{
	VARIANT var;
	VariantInit(&var);
	ptr_->get_Source(&var);
	if (var.vt==VT_DISPATCH)
	{
		command cmd(reinterpret_cast<command::interface_type*>(var.pdispVal));
		return cmd;
	}
	return command();
}

msado::string msado::recordset::get_source() const
{
	string strcmd;

	VARIANT var;
	VariantInit(&var);

	ptr_->get_Source(&var);

	if (var.vt == VT_BSTR)
	{
		strcmd = var.bstrVal;
	}

	return strcmd;
}

bool msado::recordset::add_new(
	_variant_t const& vtFieldList/* = vtMissing*/,
	_variant_t const& vtValues/* = vtMissing*/)
{
	HRESULT hr = ptr_->AddNew(vtFieldList, vtValues);
	return hr == S_OK;
}

bool msado::recordset::cancel_update()
{
	HRESULT hr=ptr_->CancelUpdate();
	return hr == S_OK;
}

bool msado::recordset::close()
{
	HRESULT hr = ptr_->Close();
	return hr == S_OK;
}

bool msado::recordset::_delete(affect AffectRecords)
{
	HRESULT hr = ptr_->Delete(static_cast<AffectEnum>(AffectRecords));
	return hr == S_OK;
}

_variant_t msado::recordset::get_rows(
	long Rows,
	const _variant_t & Start/* = vtMissing*/,
	const _variant_t & Fields/* = vtMissing*/)
{
	VARIANT var;
	VariantInit(&var);
	ptr_->GetRows(Rows, Start, Fields,&var);
	return _variant_t(var, false);
}

bool msado::recordset::move(
	long NumRecords,
	const _variant_t & Start/* = vtMissing*/)
{
	HRESULT hr =ptr_->Move(NumRecords, Start);
	return hr == S_OK;
}

bool msado::recordset::move_next()
{
	HRESULT hr = ptr_->MoveNext();
	return hr == S_OK;
}

bool msado::recordset::move_previous()
{
	HRESULT hr = ptr_->MovePrevious();
	return hr == S_OK;
}

bool msado::recordset::move_first()
{
	HRESULT hr = ptr_->MoveFirst();
	return hr == S_OK;
}

bool msado::recordset::move_last()
{
	HRESULT hr = ptr_->MoveLast();
	return hr == S_OK;
}

bool msado::recordset::open(
	const _variant_t & Source,
	const _variant_t & ActiveConnection,
	cursor_type CursorType/* = open_keyset*/,
	lock_type LockType/* = lock_optimistic*/,
	long Options/* = command_type::cmd_table*/)
{
	HRESULT hr = ptr_->Open(Source, ActiveConnection,
		static_cast<CursorTypeEnum>(CursorType),
		static_cast<LockTypeEnum>(LockType),
		Options);
	return hr == S_OK;
}

bool msado::recordset::requery(
	long Options/* = execute_option::option_unspecified*/)
{
	HRESULT hr = ptr_->Requery(Options);
	return hr == S_OK;
}

bool msado::recordset::_xresync(affect AffectRecords)
{
	HRESULT hr = ptr_->_xResync(static_cast<AffectEnum>(AffectRecords));
	return hr == S_OK;
}

bool msado::recordset::_update(
	const _variant_t & Fields/* = vtMissing*/,
	const _variant_t & Values/* = vtMissing*/)
{
	HRESULT hr=ptr_->Update(Fields, Values);
	return hr == S_OK;
}

msado::position msado::recordset::get_absolute_page() const
{
	PositionEnum pos;
	ptr_->get_AbsolutePage(&pos);
	return static_cast<position>(pos);
}

void msado::recordset::set_absolute_page(position pl/* = pos_bof*/)
{
	ptr_->put_AbsolutePage(static_cast<PositionEnum>(pl));
}

msado::edit_mode msado::recordset::get_edit_mode() const
{
	EditModeEnum em;
	ptr_->get_EditMode(&em);
	return static_cast<edit_mode>(em);
}

_variant_t msado::recordset::get_filter() const
{
	VARIANT var;
	VariantInit(&var);
	ptr_->get_Filter(&var);
	return _variant_t(var, false);
}

void msado::recordset::set_filter(const _variant_t & Criteria)
{
	ptr_->put_Filter(Criteria);
}

long msado::recordset::get_page_count() const
{
	long result = 0;
	ptr_->get_PageCount(&result);
	return result;
}

long msado::recordset::get_page_size() const
{
	long result = 0;
	ptr_->get_PageSize(&result);
	return result;
}

void msado::recordset::set_page_size(long pl)
{
	ptr_->put_PageSize(pl);
}

msado::string msado::recordset::get_sort() const
{
	BSTR bstr = NULL;
	ptr_->get_Sort(&bstr);
	return bstr;
}

void msado::recordset::set_sort(string const& Criteria)
{
	_bstr_t bstr = detail::string_to_bstr_t(Criteria);
	ptr_->put_Sort(bstr.GetBSTR());
}

long msado::recordset::get_status() const
{
	long result = 0;
	ptr_->get_Status(&result);
	return result;
}

long msado::recordset::get_state() const
{
	long result = 0;
	ptr_->get_State(&result);
	return result;
}

msado::recordset msado::recordset::_xclone()
{
	recordset set;
	ptr_->_xClone(set.get_address_of());
	return set;
}

bool msado::recordset::update_batch(affect AffectRecords)
{
	HRESULT hr = ptr_->UpdateBatch(static_cast<AffectEnum>(AffectRecords));
	return hr == S_OK;
}

bool msado::recordset::cancel_batch(affect AffectRecords)
{
	HRESULT hr = ptr_->CancelBatch(static_cast<AffectEnum>(AffectRecords));
	return hr == S_OK;
}

msado::cursor_location msado::recordset::get_cursor_location() const
{
	CursorLocationEnum cl;
	ptr_->get_CursorLocation(&cl);
	return static_cast<cursor_location>(cl);
}

void msado::recordset::set_cursor_location(cursor_location plCursorLoc)
{
	ptr_->put_CursorLocation(static_cast<CursorLocationEnum>(plCursorLoc));
}

msado::recordset msado::recordset::next_recordset(long * RecordsAffected)
{
	recordset set;
	VARIANT var;
	VariantInit(&var);
	ptr_->NextRecordset(&var, set.get_address_of());
	if (var.vt==VT_I4)
	{
		*RecordsAffected = var.lVal;
	}
	return set;
}

bool msado::recordset::supports(cursor_option CursorOptions)
{
	VARIANT_BOOL vb = 0;
	ptr_->Supports(static_cast<CursorOptionEnum>(CursorOptions), &vb);
	return vb == VARIANT_TRUE;
}

//TODO
_variant_t msado::recordset::get_collect(const _variant_t & Index) const
{
	VARIANT var;
	VariantInit(&var);
	ptr_->get_Collect(Index, &var);
	return _variant_t(var, false);
}

//TODO
void msado::recordset::set_collect(
	const _variant_t & Index,
	const _variant_t & pvar)
{
	ptr_->put_Collect(Index, pvar);
}

msado::marshal_options msado::recordset::get_marshal_options() const
{
	MarshalOptionsEnum mo;
	ptr_->get_MarshalOptions(&mo);
	return static_cast<marshal_options>(mo);
}

void msado::recordset::set_marshal_options(marshal_options peMarshal)
{
	ptr_->put_MarshalOptions(static_cast<MarshalOptionsEnum>(peMarshal));
}

bool msado::recordset::find(
	string const& Criteria,
	long SkipRecords,
	search_direction SearchDirection/* = search_forward*/,
	const _variant_t & Start/* = vtMissing*/)
{
	_bstr_t criteria = detail::string_to_bstr_t(Criteria);
	HRESULT hr = ptr_->Find(criteria.GetBSTR(), SkipRecords, 
		static_cast<SearchDirectionEnum>(SearchDirection), Start);
	return hr == S_OK;
}

bool msado::recordset::cancel()
{
	HRESULT hr=ptr_->Cancel();
	return hr == S_OK;
}

IUnknownPtr msado::recordset::get_data_source() const
{
	IUnknownPtr puk;
	ptr_->get_DataSource(&puk);
	return puk;
}

void msado::recordset::set_ref_data_source(IUnknown * ppunkDataSource)
{
	ptr_->putref_DataSource(ppunkDataSource);
}

bool msado::recordset::_xsave(
	string const& FileName,
	persist_format PersistFormat)
{
	_bstr_t bstr = detail::string_to_bstr_t(FileName);
	HRESULT hr=ptr_->_xSave(bstr.GetBSTR(),
		static_cast<PersistFormatEnum>(PersistFormat));
	return hr == S_OK;
}

msado::command msado::recordset::get_active_command() const
{
	command cmd;
	ptr_->get_ActiveCommand((IDispatch**)cmd.get_address_of());
	return cmd;
}

void msado::recordset::set_stay_in_sync(bool pbStayInSync)
{
	ptr_->put_StayInSync(pbStayInSync ? VARIANT_TRUE : VARIANT_FALSE);
}

bool msado::recordset::get_stay_in_sync() const
{
	VARIANT_BOOL vb = 0;
	ptr_->get_StayInSync(&vb);
	return vb == VARIANT_TRUE;
}

msado::string msado::recordset::get_string(
	string_format StringFormat, //=clip_string
	long NumRows,
	string const& ColumnDelimeter,
	string const& RowDelimeter,
	string const& NullExpr) const
{
	BSTR bstr = NULL;
	ptr_->GetString(static_cast<StringFormatEnum>(StringFormat),
		NumRows,
		detail::string_to_bstr_t(ColumnDelimeter).GetBSTR(),
		detail::string_to_bstr_t(RowDelimeter).GetBSTR(),
		detail::string_to_bstr_t(NullExpr).GetBSTR(),&bstr);
	return bstr;
}

msado::string msado::recordset::get_data_member() const
{
	BSTR bstr = NULL;
	ptr_->get_DataMember(&bstr);
	return bstr;
}

void msado::recordset::set_data_member(string const& pbstrDataMember)
{
	_bstr_t bstr = detail::string_to_bstr_t(pbstrDataMember);
	ptr_->put_DataMember(bstr.GetBSTR());
}

msado::_compare msado::recordset::compare_bookmarks(
	const _variant_t & Bookmark1,
	const _variant_t & Bookmark2)
{
	CompareEnum cp;
	ptr_->CompareBookmarks(Bookmark1, Bookmark2, &cp);
	return static_cast<_compare>(cp);
}

msado::recordset msado::recordset::clone(lock_type LockType)
{
	recordset set;
	ptr_->Clone(static_cast<LockTypeEnum>(LockType), set.get_address_of());
	return set;
}

bool msado::recordset::resync(
	affect AffectRecords,
	_resync ResyncValues)
{
	HRESULT hr = ptr_->Resync(
		static_cast<AffectEnum>(AffectRecords),
		static_cast<ResyncEnum>(ResyncValues));
	return hr == S_OK;
}

bool msado::recordset::seek(
	const _variant_t & KeyValues,
	_seek SeekOption/*=seek_first_eq*/)
{
	HRESULT hr=ptr_->Seek(KeyValues, static_cast<SeekEnum>(SeekOption));
	return hr == S_OK;
}

void msado::recordset::set_index(string const& pbstrIndex)
{
	_bstr_t index = detail::string_to_bstr_t(pbstrIndex);
	ptr_->put_Index(index.GetBSTR());
}

msado::string msado::recordset::get_index() const
{
	BSTR bstr = NULL;
	HRESULT hr=ptr_->get_Index(&bstr);
	return bstr;
}

bool msado::recordset::save(
	const _variant_t & Destination,
	persist_format PersistFormat/* = persist_adtg*/)
{
	HRESULT hr=ptr_->Save(Destination,
		static_cast<PersistFormatEnum>(PersistFormat));
	return hr == S_OK;
}

///////////////////////////////////////////////////////
//

void msado::recordset::on_field_changing(
	long cFields, _variant_t const& Fields)
{

}

void msado::recordset::on_field_changed(
	long cFields, _variant_t const& Fields, error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::recordset::on_record_changing(event_reason adReason, long cRecords)
{

}

void msado::recordset::on_record_changed(
	event_reason adReason, long cRecords, error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::recordset::on_recordset_changing(
	event_reason adReason)
{

}

void msado::recordset::on_recordset_changed(
	event_reason adReason, error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::recordset::on_moving(event_reason adReason)
{

}

void msado::recordset::on_moved(
	event_reason adReason, error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}

void msado::recordset::on_recordset_end(bool fMoreData)
{

}

void msado::recordset::on_fetching(
	long Progress, long MaxProgress)
{

}

void msado::recordset::on_fetched(error const& e)
{
	if (e)
	{
		//msado::string desc = e.get_description();
	}
}
