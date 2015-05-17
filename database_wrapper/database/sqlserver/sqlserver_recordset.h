#pragma once

namespace database
{
	//
	//表示用于 SQL SERVER 的结果集操作
	//
	template<typename DataBase,
		template<typename>class SqlService/* = sqlserver_recordset_service*/>
	class sqlserver_recordset 
		: public detail::basic_recordset<DataBase, SqlService>
	{
	public:
		explicit sqlserver_recordset(
			CLSID const& clsid = msado::Clsid::ADODB_Recordset())
			: detail::basic_recordset<DataBase, SqlService>(clsid)
		{
		}

		sqlserver_recordset(DataBase const& db,msado::recordset& set)
			: detail::basic_recordset<DataBase, SqlService>(db,set)
		{

		}
	public:
		bool open(string const& table, sqlserver_connect<DataBase>& conn,
			msado::cursor_type cursortype = msado::open_keyset,
			msado::lock_type locktype = msado::lock_optimistic,
			long options = msado::cmd_table);

		string get_string_value(long index) const;
		string get_string_value(string const& index) const;

		long get_long_value(long index) const;
		long get_long_value(string const& index) const;

		unsigned long get_ulong_value(long index) const;
		unsigned long get_ulong_value(string const& index) const;

		double get_double_value(long index) const;
		double get_double_value(string const& index) const;

		datetime get_datetime_value(long index) const;
		datetime get_datetime_value(string const& index) const;

		bool set_value(long index, string const& val);
		bool set_value(string const& index, string const& val);

		bool set_value(long index, long val);
		bool set_value(string const& index, long val);

		bool set_value(long index, unsigned long val);
		bool set_value(string const& index, unsigned long val);

		bool set_value(long index, double val);
		bool set_value(string const& index, double val);

		bool set_value(long index, datetime const& val);
		bool set_value(string const& index, datetime const& val);

		bool move(long numrecords);
		bool move_next();
		bool move_previous();
		bool move_first();
		bool move_last();

		bool find(string const& criteria, long skiprecords,
			msado::search_direction searchdirection = msado::search_forward);

		bool cancel();

		bool get_bof() const;
		bool get_eof() const;

		long get_record_count() const;

		bool add_new();
		bool add_new(long index, string const& val);
		bool add_new(string const& index, string const& val);
		bool add_new(long index, long val);
		bool add_new(string const& index, long val);
		bool add_new(long index, unsigned long val);
		bool add_new(string const& index, unsigned long val);
		bool add_new(long index, double val);
		bool add_new(string const& index, double val);
		bool add_new(long index, datetime const& val);
		bool add_new(string const& index, datetime const& val);

		bool update();
		bool update_batch(msado::affect affectrecords);
		bool cancel_update();
		bool cancel_batch(msado::affect affectrecords);

	public:

		void set_ref_active_connection(sqlserver_connect<DataBase>& pvar);
		sqlserver_connect<DataBase> get_ref_active_connection() const;
		void set_active_connection(string const& connstr);
		string get_active_connection() const;
		_variant_t get_bookmark() const;
		void set_bookmark(_variant_t const& pvbookmark);
		long get_cache_size() const;
		void set_cache_size(long pl);
		msado::cursor_type get_cursor_type() const;
		void set_cursor_type(msado::cursor_type plcursortype = msado::open_dynamic);
		msado::lock_type get_lock_type() const;
		void set_lock_type(msado::lock_type pllocktype = msado::lock_read_only);
		long get_max_records() const;
		void set_max_records(long plmaxrecords);
		void set_ref_source(sqlserver_command<DataBase>& cmd);
		void set_source(string const& pvsource);
		sqlserver_command<DataBase> get_ref_source() const;
		string get_source() const;
		bool _delete(msado::affect affectrecords);
		//safearray
		//_variant_t get_rows(long Rows) const;
		bool requery(long options = msado::/*execute_option::*/option_unspecified);
		bool _xresync(msado::affect affectrecords);
		msado::position get_absolute_page() const;
		void set_absolute_page(msado::position pl = msado::pos_bof);
		msado::edit_mode get_edit_mode() const;
		//_variant_t get_filter() const;
		string get_condition_filter() const;
		msado::filter_group get_group_filter() const;
		//void set_filter(const _variant_t & criteria);
		void set_condition_filter(string const& condition);
		void set_group_filter(msado::filter_group groupfilter);
		long get_page_count() const;
		long get_page_size() const;
		void set_page_size(long pl);
		string get_sort() const;
		void set_sort(string const& criteria);
		long get_status() const;
		long get_state() const;
		sqlserver_recordset<DataBase,SqlService> _xclone();
		msado::cursor_location get_cursor_location() const;
		void set_cursor_location(msado::cursor_location plcursorloc);
		sqlserver_recordset<DataBase,SqlService> next_recordset(
			long * recordsaffected);
		bool supports(msado::cursor_option cursoroptions);
		string get_collect(long index) const;
		string get_collect(string const& index) const;
		void set_collect(long index, string const& pvar);
		void set_collect(string const& index, string const& pvar);
		msado::marshal_options get_marshal_options() const;
		void set_marshal_options(msado::marshal_options pemarshal);
		bool _xsave(string const& filename, msado::persist_format persistformat);
		sqlserver_command<DataBase> get_active_command() const;
		void set_stay_in_sync(bool pbstayinsync);
		bool get_stay_in_sync() const;
		string get_string(msado::string_format stringformat,
			long numrows,
			string const& columndelimeter,
			string const& rowdelimeter,
			string const& _nullexpr) const;
		string get_data_member() const;
		void set_data_member(string const& pbstrdatamember);
		//msado::_compare compare_bookmarks(
		//	_variant_t const& bookmark1, _variant_t const& bookmark2);
		sqlserver_recordset<DataBase,SqlService> clone(msado::lock_type locktype);
		bool resync(
			msado::affect affectrecords, msado::_resync resyncvalues);
		//bool seek(
		//	_variant_t const& keyvalues, msado::_seek seekoption = seek_first_eq);
		//void set_index(string const& pbstrindex);
		//string get_index() const;
		//bool save(_variant_t const& destination,
		//	msado::persist_format persistformat = msado::persist_adtg);
	public:
		typename implement_type::interface_type const* get_internal_ptr() const;
		typename implement_type::interface_type * get_internal_ptr();
	};
}

namespace database
{
	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::open(
		string const& table, sqlserver_connect<DataBase>& conn,
		msado::cursor_type cursortype/* = msado::open_keyset*/,
		msado::lock_type locktype/* = msado::lock_optimistic*/,
		long options/* = msado::cmd_table*/)
	{
		return get_service().open(get_implement(),table,conn,
			cursortype,locktype,options);
	}

	template<typename DataBase,template<typename>class SqlService>
	inline string sqlserver_recordset<DataBase, SqlService>::get_string_value(
		long index) const
	{
		return get_service().get_string_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline string sqlserver_recordset<DataBase, SqlService>::get_string_value(
		string const& index) const
	{
		return get_service().get_string_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_long_value(
		long index) const
	{
		return get_service().get_long_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_long_value(
		string const& index) const
	{
		return get_service().get_long_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline unsigned long sqlserver_recordset<DataBase, SqlService>::get_ulong_value(
		long index) const
	{
		return get_service().get_ulong_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline unsigned long sqlserver_recordset<DataBase, SqlService>::get_ulong_value(
		string const& index) const
	{
		return get_service().get_ulong_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline double sqlserver_recordset<DataBase, SqlService>::get_double_value(
		long index) const
	{
		return get_service().get_double_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline double sqlserver_recordset<DataBase, SqlService>::get_double_value(
		string const& index) const
	{
		return get_service().get_double_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline datetime sqlserver_recordset<DataBase, SqlService>::get_datetime_value(
		long index) const
	{
		return get_service().get_datetime_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline datetime sqlserver_recordset<DataBase, SqlService>::get_datetime_value(
		string const& index) const
	{
		return get_service().get_datetime_value(get_implement(), index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		long index, string const& val)
	{
		return get_service().set_value(get_implement(), index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		string const& index, string const& val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		long index, long val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		string const& index, long val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		long index, unsigned long val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		string const& index, unsigned long val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		long index, double val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		string const& index, double val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		long index, datetime const& val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::set_value(
		string const& index, datetime const& val)
	{
		return get_service().set_value(get_implement(), index, val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::move(long numrecords)
	{
		return get_service().move(get_implement(), numrecords);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::move_next()
	{
		return get_service().move_next(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::move_previous()
	{
		return get_service().move_previous(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::move_first()
	{
		return get_service().move_first(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::move_last()
	{
		return get_service().move_last(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::find(
		string const& criteria, long skiprecords,
		msado::search_direction searchdirection/* = msado::search_forward*/)
	{
		return get_service().find(get_implement(),
			criteria,skiprecords,searchdirection);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::cancel()
	{
		return get_service().cancel(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::get_bof() const
	{
		return get_service().get_bof(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::get_eof() const
	{
		return get_service().get_eof(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_record_count() const
	{
		return get_service().get_record_count(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new()
	{
		return get_service().add_new(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		long index, string const& val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		string const& index, string const& val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		long index, long val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		string const& index, long val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		long index, unsigned long val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		string const& index, unsigned long val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		long index, double val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		string const& index, double val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		long index, datetime const& val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::add_new(
		string const& index, datetime const& val)
	{
		return get_service().add_new(get_implement(),index,val);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::update()
	{
		return get_service().update(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::update_batch(
		msado::affect affectrecords)
	{
		return get_service().update_batch(get_implement(),affectrecords);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::cancel_update()
	{
		return get_service().cancel_update(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::cancel_batch(
		msado::affect affectrecords)
	{
		return get_service().cancel_batch(get_implement(),affectrecords);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void 
		sqlserver_recordset<DataBase, SqlService>::set_ref_active_connection(
			sqlserver_connect<DataBase>& pvar)
	{
		get_service().set_ref_active_connection(get_implement(),pvar);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_connect<DataBase> 
		sqlserver_recordset<DataBase, SqlService>::get_ref_active_connection() const
	{
		return get_service().get_ref_active_connection(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_active_connection(
		string const& connstr)
	{
		get_service().set_active_connection(get_implement(),connstr);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline string 
		sqlserver_recordset<DataBase, SqlService>::get_active_connection() const
	{
		return get_service().get_active_connection(get_implement());
	}

	//_variant_t get_bookmark() const;
	//void set_bookmark(_variant_t const& pvbookmark);

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_cache_size() const
	{
		return get_service().get_cache_size(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_cache_size(long pl)
	{
		get_service().set_cache_size(get_implement(),pl);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::cursor_type 
		sqlserver_recordset<DataBase, SqlService>::get_cursor_type() const
	{
		return get_service().get_cursor_type(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_cursor_type(
		msado::cursor_type plcursortype/* = msado::open_dynamic*/)
	{
		get_service().set_cursor_type(get_implement(),plcursortype);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::lock_type 
		sqlserver_recordset<DataBase, SqlService>::get_lock_type() const
	{
		return get_service().get_lock_type(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_lock_type(
		msado::lock_type pllocktype/* = msado::lock_read_only*/)
	{
		get_service().set_lock_type(get_implement(),pllocktype);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_max_records() const
	{
		return get_service().get_max_records(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_max_records(
		long plmaxrecords)
	{
		get_service().set_max_records(get_implement(),plmaxrecords);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_ref_source(
		sqlserver_command<DataBase>& cmd)
	{
		get_service().set_ref_source(get_implement(),cmd);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_source(
		string const& pvsource)
	{
		get_service().set_source(get_implement(),pvsource);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_command<DataBase> 
		sqlserver_recordset<DataBase, SqlService>::get_ref_source() const
	{
		return get_service().get_ref_source(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline string sqlserver_recordset<DataBase, SqlService>::get_source() const
	{
		return get_service().get_source(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::_delete(
		msado::affect affectrecords)
	{
		return get_service()._delete(get_implement(),affectrecords);
	}

	//safearray
	//_variant_t get_rows(long Rows) const;

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::requery(
		long options /*= msado::execute_option::option_unspecified*/)
	{
		return get_service().requery(get_implement(),options);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::_xresync(
		msado::affect affectrecords)
	{
		return get_service()._xresync(get_implement(),affectrecords);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::position 
		sqlserver_recordset<DataBase, SqlService>::get_absolute_page() const
	{
		return get_service().get_absolute_page(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_absolute_page(
		msado::position pl/* = msado::pos_bof*/)
	{
		get_service().set_absolute_page(get_implement(),pl);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::edit_mode 
		sqlserver_recordset<DataBase, SqlService>::get_edit_mode() const
	{
		return get_service().get_edit_mode(get_implement());
	}

	//_variant_t get_filter() const;

	template<typename DataBase, template<typename>class SqlService>
	inline database::string 
		sqlserver_recordset<DataBase, SqlService>::get_condition_filter() const
	{
		return get_service().get_condition_filter(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::filter_group 
		sqlserver_recordset<DataBase, SqlService>::get_group_filter() const
	{
		return get_service().get_group_filter(get_implement());
	}

	//void set_filter(const _variant_t & criteria);

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_condition_filter(
		string const& condition)
	{
		get_service().set_condition_filter(get_implement(),condition);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_group_filter(
		msado::filter_group groupfilter)
	{
		get_service().set_group_filter(get_implement(),groupfilter);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_page_count() const
	{
		return get_service().get_page_count(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_page_size() const
	{
		return get_service().get_page_size(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_page_size(long pl)
	{
		get_service().set_page_size(get_implement(),pl);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline database::string 
		sqlserver_recordset<DataBase, SqlService>::get_sort() const
	{
		return get_service().get_sort(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_sort(
		string const& criteria)
	{
		get_service().set_sort(get_implement(),criteria);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_status() const
	{
		return get_service().get_status(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline long sqlserver_recordset<DataBase, SqlService>::get_state() const
	{
		return get_service().get_state(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_recordset<DataBase,SqlService> 
		sqlserver_recordset<DataBase, SqlService>::_xclone()
	{
		return get_service()._xclone(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::cursor_location 
		sqlserver_recordset<DataBase, SqlService>::get_cursor_location() const
	{
		return get_service().get_cursor_location(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_cursor_location(
		msado::cursor_location plcursorloc)
	{
		get_service().set_cursor_location(get_implement(),plcursorloc);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_recordset<DataBase,SqlService> 
		sqlserver_recordset<DataBase, SqlService>::next_recordset(
			long * recordsaffected)
	{
		return get_service().next_recordset(get_implement(),recordsaffected);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::supports(
		msado::cursor_option cursoroptions)
	{
		return get_service().supports(get_implement(),cursoroptions);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline database::string 
		sqlserver_recordset<DataBase, SqlService>::get_collect(long index) const
	{
		return get_service().get_collect(get_implement(),index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline database::string 
		sqlserver_recordset<DataBase, SqlService>::get_collect(
			string const& index) const
	{
		return get_service().get_collect(get_implement(),index);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_collect(
		long index, string const& pvar)
	{
		get_service().set_collect(get_implement(),index,pvar);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_collect(
		string const& index, string const& pvar)
	{
		get_service().set_collect(get_implement(),index,pvar);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline msado::marshal_options 
		sqlserver_recordset<DataBase, SqlService>::get_marshal_options() const
	{
		return get_service().get_marshal_options(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_marshal_options(
		msado::marshal_options pemarshal)
	{
		get_service().set_marshal_options(get_implement(),pemarshal);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::_xsave(
		string const& filename, msado::persist_format persistformat)
	{
		return get_service()._xsave(get_implement(),filename,persistformat);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_command<DataBase> 
		sqlserver_recordset<DataBase, SqlService>::get_active_command() const
	{
		return get_service().get_active_command(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_stay_in_sync(
		bool pbstayinsync)
	{
		get_service().set_stay_in_sync(get_implement(),pbstayinsync);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::get_stay_in_sync() const
	{
		return get_service().get_stay_in_sync(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline database::string
		sqlserver_recordset<DataBase, SqlService>::get_string(
			msado::string_format stringformat,
			long numrows,
			string const& columndelimeter,
			string const& rowdelimeter,
			string const& _nullexpr) const
	{
		return get_service().get_string(get_implement(), stringformat,
			numrows, columndelimeter, rowdelimeter, _nullexpr);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline database::string 
		sqlserver_recordset<DataBase, SqlService>::get_data_member() const
	{
		return get_service().get_data_member(get_implement());
	}

	template<typename DataBase, template<typename>class SqlService>
	inline void sqlserver_recordset<DataBase, SqlService>::set_data_member(
		string const& pbstrdatamember)
	{
		get_service().set_data_memeber(get_implement(),pbstrdatamember);
	}

	//msado::_compare compare_bookmarks(
	//	_variant_t const& bookmark1, _variant_t const& bookmark2);

	template<typename DataBase, template<typename>class SqlService>
	inline sqlserver_recordset<DataBase,SqlService> 
		sqlserver_recordset<DataBase, SqlService>::clone(msado::lock_type locktype)
	{
		return get_service().clone(get_implement(),locktype);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline bool sqlserver_recordset<DataBase, SqlService>::resync(
		msado::affect affectrecords, msado::_resync resyncvalues)
	{
		return get_service().resync(get_implement(),affectrecords,resyncvalues);
	}

	template<typename DataBase, template<typename>class SqlService>
	inline typename sqlserver_recordset<DataBase, SqlService>::
		implement_type::interface_type const*
		sqlserver_recordset<DataBase, SqlService>::get_internal_ptr() const
	{
		return get_implement().get();
	}
	
	template<typename DataBase, template<typename>class SqlService>
	inline typename sqlserver_recordset<DataBase, SqlService>::
		implement_type::interface_type *
		sqlserver_recordset<DataBase, SqlService>::get_internal_ptr()
	{
		return get_implement().get();
	}
}