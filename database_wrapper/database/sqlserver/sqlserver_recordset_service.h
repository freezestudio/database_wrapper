#pragma once

namespace database
{
	template<typename DataBase>
	class sqlserver_recordset_service
		: public detail::basic_service<sqlserver_recordset_service<DataBase> >
	{
	public:
		typedef msado::recordset implement_type;
		typedef DataBase database_type;
	public:
		sqlserver_recordset_service()
		{

		}

		explicit sqlserver_recordset_service(database_type const& db)
			: database_(db)
		{

		}

		sqlserver_recordset_service(sqlserver_recordset_service const& rhs)
			: database_(rhs.database_)
		{

		}
	//public:
	//	database_type get_database() const
	//	{
	//		return database_;
	//	}
	public:
		bool open(implement_type& impl,
			string const& _select, sqlserver_connect<DataBase>& conn);

		bool close(implement_type& impl);

		string get_string_value(implement_type const& impl, long index) const;
		string get_string_value(implement_type const& impl, string const& index) const;

		long get_long_value(implement_type const& impl, long index) const;
		long get_long_value(implement_type const& impl, string const& index) const;

		unsigned long get_ulong_value(implement_type const& impl, long index) const;
		unsigned long get_ulong_value(implement_type const& impl, string const& index) const;

		double get_double_value(implement_type const& impl, long index) const;
		double get_double_value(implement_type const& impl, string const& index) const;

		datetime get_datetime_value(implement_type const& impl, long index) const;
		datetime get_datetime_value(implement_type const& impl, string const& index) const;

		bool set_value(implement_type& impl, long index,string const& val);
		bool set_value(implement_type& impl, string const& index, string const& val);

		bool set_value(implement_type& impl, long index, long val);
		bool set_value(implement_type& impl, string const& index, long val);

		bool set_value(implement_type& impl, long index,unsigned long val);
		bool set_value(implement_type& impl, string const& index, unsigned long val);

		bool set_value(implement_type& impl, long index, double val);
		bool set_value(implement_type& impl, string const& index, double val);

		bool set_value(implement_type& impl, long index, datetime const& val);
		bool set_value(implement_type& impl, string const& index, datetime const& val);
		
		bool move(implement_type& impl, long numrecords);
		bool move_next(implement_type& impl );
		bool move_previous(implement_type& impl );
		bool move_first(implement_type& impl );
		bool move_last(implement_type& impl );

		bool find(implement_type& impl,
			string const& criteria,long skiprecords,
			msado::search_direction searchdirection = msado::search_forward);

		bool cancel(implement_type& impl );

		bool get_bof(implement_type const& impl ) const;
		bool get_eof(implement_type const& impl ) const;

		long get_record_count(implement_type const& impl ) const;

		bool add_new(implement_type& impl);
		bool add_new(implement_type& impl, long index, string const& val);
		bool add_new(implement_type& impl, string const& index, string const& val);
		bool add_new(implement_type& impl, long index, long val);
		bool add_new(implement_type& impl, string const& index, long val);
		bool add_new(implement_type& impl, long index, unsigned long val);
		bool add_new(implement_type& impl, string const& index, unsigned long val);
		bool add_new(implement_type& impl, long index, double val);
		bool add_new(implement_type& impl, string const& index, double val);
		bool add_new(implement_type& impl, long index, datetime const& val);
		bool add_new(implement_type& impl, string const& index, datetime const& val);

		bool update(implement_type& impl);
		bool update_batch(implement_type& impl, msado::affect affectrecords);
		bool cancel_update(implement_type& impl );
		bool cancel_batch(implement_type& impl, msado::affect affectrecords);

	public:

		void set_ref_active_connection(implement_type& impl, 
			sqlserver_connect<DataBase>& pvar);
		sqlserver_connect<DataBase> get_ref_active_connection(
			implement_type const& impl) const;
		void set_active_connection(implement_type& impl, string const& connstr);
		string get_active_connection(implement_type const& impl) const;
		_variant_t get_bookmark(implement_type const& impl) const;
		void set_bookmark(implement_type& impl, _variant_t const& pvbookmark);
		long get_cache_size(implement_type const& impl) const;
		void set_cache_size(implement_type& impl, long pl);
		msado::cursor_type get_cursor_type(implement_type const& impl) const;
		void set_cursor_type(implement_type& impl,
			msado::cursor_type plcursortype = msado::open_dynamic);
		msado::lock_type get_lock_type(implement_type const& impl) const;
		void set_lock_type(implement_type& impl,
			msado::lock_type pllocktype = msado::lock_read_only);
		long get_max_records(implement_type const& impl) const;
		void set_max_records(implement_type& impl, long plmaxrecords);
		void set_ref_source(implement_type& impl, sqlserver_command<DataBase>& cmd);
		void set_source(implement_type& impl, string const& pvsource);
		sqlserver_command<DataBase> get_ref_source(implement_type const& impl) const;
		string get_source(implement_type const& impl) const;
		bool _delete(implement_type& impl, msado::affect affectrecords);
		//safearray
		/*unused*/_variant_t get_rows(implement_type const& impl,long Rows) const;
		bool requery(implement_type& impl,
			long options = msado::/*execute_option::*/option_unspecified);
		bool _xresync(implement_type& impl, msado::affect affectrecords);
		msado::position get_absolute_page(implement_type const& impl) const;
		void set_absolute_page(implement_type& impl,
			msado::position pl = msado::pos_bof);
		msado::edit_mode get_edit_mode(implement_type const& impl) const;
		/*unused*/_variant_t get_filter(implement_type const& impl) const;
		string get_condition_filter(implement_type const& impl) const;
		msado::filter_group get_group_filter(implement_type const& impl) const;
		/*unused*/void set_filter(implement_type& impl, const _variant_t & criteria);
		void set_condition_filter(implement_type& impl, string const& condition);
		void set_group_filter(implement_type& impl, msado::filter_group groupfilter);
		long get_page_count(implement_type const& impl) const;
		long get_page_size(implement_type const& impl) const;
		void set_page_size(implement_type& impl, long pl);
		string get_sort(implement_type const& impl) const;
		void set_sort(implement_type& impl, string const& criteria);
		long get_status(implement_type const& impl) const;
		long get_state(implement_type const& impl) const;
		sqlserver_recordset<DataBase> _xclone(implement_type& impl );
		msado::cursor_location get_cursor_location(implement_type const& impl) const;
		void set_cursor_location(implement_type& impl,
			msado::cursor_location plcursorloc);
		sqlserver_recordset<DataBase> next_recordset(implement_type& impl,
			long * recordsaffected);
		bool supports(implement_type& impl, msado::cursor_option cursoroptions);
		string get_collect(implement_type const& impl,long index) const;
		string get_collect(implement_type const& impl, string const& index) const;
		void set_collect(implement_type& impl,long index,string const& pvar);
		void set_collect(implement_type& impl,
			string const& index, string const& pvar);
		msado::marshal_options get_marshal_options(implement_type const& impl) const;
		void set_marshal_options(implement_type& impl,
			msado::marshal_options pemarshal);
		bool _xsave(implement_type& impl,
			string const& filename,msado::persist_format persistformat);
		sqlserver_command<DataBase> get_active_command(implement_type const& impl) const;
		void set_stay_in_sync(implement_type& impl, bool pbstayinsync);
		bool get_stay_in_sync(implement_type const& impl) const;
		string get_string(implement_type const& impl,
			msado::string_format stringformat,long numrows,
			string const& columndelimeter,
			string const& rowdelimeter,
			string const& _nullexpr) const;
		string get_data_member(implement_type const& impl) const;
		void set_data_member(implement_type& impl, string const& pbstrdatamember);
		/*unused*/msado::_compare compare_bookmarks(implement_type& impl,
			 _variant_t const& bookmark1, _variant_t const& bookmark2);
		sqlserver_recordset<DataBase> clone(
			implement_type& impl, msado::lock_type locktype);
		bool resync(implement_type& impl,
			msado::affect affectrecords,msado::_resync resyncvalues);
		/*unused*/bool seek(implement_type& impl,
			 _variant_t const& keyvalues,msado::_seek seekoption = seek_first_eq);
		/*unused*/void set_index(implement_type& impl,string const& pbstrindex);
		/*unused*/string get_index(implement_type const& impl) const;
		/*unused*/bool save(implement_type& impl,	_variant_t const& destination,
			msado::persist_format persistformat = msado::persist_adtg);
	private:
		bool opened(implement_type const& impl) const;
		_variant_t get_value_impl(implement_type const& impl, long index) const;
		_variant_t get_value_impl(implement_type const& impl, string const& index) const;
		bool set_value_impl(implement_type& impl, long index, _variant_t const& val);
		bool set_value_impl(implement_type& impl, string const& index, _variant_t const& val);
	private:
		database_type database_;
	};
}

namespace database
{
	//template<typename DataBase>
	//inline bool sqlserver_recordset_service<DataBase>::open(implement_type& impl)
	//{
	//	return false;
	//}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::open(implement_type& impl,
		string const& _select, sqlserver_connect<DataBase>& conn)
	{
		return impl.open(_variant_t(static_cast<const wchar_t*>(_select)),
			_variant_t(conn.get_internal_ptr()/*, false*/), msado::open_static);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::close(implement_type& impl)
	{
		return impl.close();
	}

	template<typename DataBase>
	inline string sqlserver_recordset_service<DataBase>::get_string_value(
		implement_type const& impl, long index) const
	{
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_BSTR) ||
			var.vt == static_cast<VARTYPE>(VT_LPWSTR))
		{
			_bstr_t bstr = static_cast<_bstr_t>(var);
			return bstr.GetBSTR();
		}

		return L"";
	}

	template<typename DataBase>
	inline string sqlserver_recordset_service<DataBase>::get_string_value(
		implement_type const& impl, string const& index) const
	{
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_BSTR) ||
			var.vt == static_cast<VARTYPE>(VT_LPWSTR))
		{
			_bstr_t bstr = static_cast<_bstr_t>(var);
			return bstr.GetBSTR();
		}

		return L"";
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_long_value(
		implement_type const& impl, long index) const
	{
		long result = 0;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_I2) ||
			var.vt == static_cast<VARTYPE>(VT_I4) ||
			var.vt == static_cast<VARTYPE>(VT_BOOL) ||
			var.vt == static_cast<VARTYPE>(VT_I8) ||
			var.vt == static_cast<VARTYPE>(VT_INT) )
		{
			result = static_cast<long>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_long_value(
		implement_type const& impl, string const& index) const
	{
		long result = 0;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_I2) ||
			var.vt == static_cast<VARTYPE>(VT_I4) ||
			var.vt == static_cast<VARTYPE>(VT_BOOL) ||
			var.vt == static_cast<VARTYPE>(VT_I8) ||
			var.vt == static_cast<VARTYPE>(VT_INT))
		{
			result = static_cast<long>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline unsigned long sqlserver_recordset_service<DataBase>::get_ulong_value(
		implement_type const& impl, long index) const
	{
		unsigned long result = 0;
		_variant_t var = get_value_impl(impl, index);

		if (
			var.vt == static_cast<VARTYPE>(VT_UI2) ||
			var.vt == static_cast<VARTYPE>(VT_UI4) ||
			var.vt == static_cast<VARTYPE>(VT_UI8) ||
			var.vt == static_cast<VARTYPE>(VT_UINT) )
		{
			result = static_cast<unsigned long>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline unsigned long sqlserver_recordset_service<DataBase>::get_ulong_value(
		implement_type const& impl, string const& index) const
	{
		long result = 0;
		_variant_t var = get_value_impl(impl, index);

		if (
			var.vt == static_cast<VARTYPE>(VT_UI2) ||
			var.vt == static_cast<VARTYPE>(VT_UI4) ||
			var.vt == static_cast<VARTYPE>(VT_UI8) ||
			var.vt == static_cast<VARTYPE>(VT_UINT))
		{
			result = static_cast<unsigned long>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline double sqlserver_recordset_service<DataBase>::get_double_value(
		implement_type const& impl, long index) const
	{
		double result = 0.0;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_R4) ||
			var.vt == static_cast<VARTYPE>(VT_R8) ||
			var.vt == static_cast<VARTYPE>(VT_DECIMAL))
		{
			result = static_cast<double>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline double sqlserver_recordset_service<DataBase>::get_double_value(
		implement_type const& impl, string const& index) const
	{
		double result = 0.0;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_R4) ||
			var.vt == static_cast<VARTYPE>(VT_R8) ||
			var.vt == static_cast<VARTYPE>(VT_DECIMAL))
		{
			result = static_cast<double>(var);
		}

		return result;
	}

	template<typename DataBase>
	inline datetime sqlserver_recordset_service<DataBase>::get_datetime_value(
		implement_type const& impl, long index) const
	{
		datetime dt;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_DATE))
		{
			dt = var.date;
		}
		else if (var.vt == static_cast<VARTYPE>(VT_BSTR))
		{
			dt.ParseDateTime(var.bstrVal);
		}

		return dt;
	}

	template<typename DataBase>
	inline datetime sqlserver_recordset_service<DataBase>::get_datetime_value(
		implement_type const& impl, string const& index) const
	{
		datetime dt;
		_variant_t var = get_value_impl(impl, index);

		if (var.vt == static_cast<VARTYPE>(VT_DATE))
		{
			dt = var.date;
		}

		return dt;
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, long index, string const& val)
	{
		_variant_t var = static_cast<const wchar_t*>(val);
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, string const& index, string const& val)
	{
		_variant_t var = static_cast<const wchar_t*>(val);
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, long index, long val)
	{
		_variant_t var(static_cast<long>(val), VT_I4);
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, string const& index, long val)
	{
		_variant_t var(static_cast<long>(val), VT_I4);
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, long index, unsigned long val)
	{
		_variant_t var(static_cast<unsigned long>(val));
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, string const& index, unsigned long val)
	{
		_variant_t var(static_cast<unsigned long>(val));
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, long index, double val)
	{
		_variant_t var(static_cast<double>(val));
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, string const& index, double val)
	{
		_variant_t var(static_cast<double>(val));
		return set_value_impl(impl, index, var);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, long index, datetime const& val)
	{
		VARIANT var;
		VariantInit(&var);
		var.vt = static_cast<VARTYPE>(::VT_DATE);
		var.date = val;

		_variant_t var_t(var);
		return set_value_impl(impl, index, var_t);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value(
		implement_type& impl, string const& index, datetime const& val)
	{
		VARIANT var;
		VariantInit(&var);
		var.vt = static_cast<VARTYPE>(::VT_DATE);
		var.date = val;

		_variant_t var_t(var);
		return set_value_impl(impl, index, var_t);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::move(
		implement_type& impl, long numrecords)
	{
		return impl.move(static_cast<long>(numrecords));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::move_next(
		implement_type& impl)
	{
		return impl.move_next();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::move_previous(
		implement_type& impl)
	{
		return impl.move_previous();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::move_first(
		implement_type& impl)
	{
		return impl.move_first();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::move_last(
		implement_type& impl)
	{
		return impl.move_last();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::find(implement_type& impl,
		string const& criteria, long skiprecords,
		msado::search_direction searchdirection/* = msado::search_forward*/)
	{
		return impl.find(criteria, static_cast<long>(skiprecords),
			searchdirection);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::cancel(
		implement_type& impl)
	{
		return impl.cancel();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::get_bof(
		implement_type const& impl) const
	{
		return impl.get_BOF();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::get_eof(
		implement_type const& impl) const
	{
		return impl.get_EOF();
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_record_count(
		implement_type const& impl) const
	{
		return impl.get_record_count();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl)

	{
		return impl.add_new();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, long index, string const& val)
	{
		return impl.add_new(_variant_t < static_cast<long>(index),
			_variant_t(static_cast<const wchar_t*>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, string const& index, string const& val)

	{
		return impl.add_new(_variant_t < static_cast<const wchar_t*>(index),
			_variant_t(static_cast<const wchar_t*>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, long index, long val)
	{
		return impl.add_new(_variant_t < static_cast<long>(index),
			_variant_t(static_cast<long>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, string const& index, long val)
	{
		return impl.add_new(_variant_t < static_cast<const wchar_t*>(index),
			_variant_t(static_cast<long>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, long index, unsigned long val)
	{
		return impl.add_new(_variant_t < static_cast<long>(index),
			_variant_t(static_cast<unsigned long>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, string const& index, unsigned long val)
	{
		return impl.add_new(_variant_t < static_cast<const wchar_t*>(index),
			_variant_t(static_cast<unsigned long>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, long index, double val)
	{
		return impl.add_new(_variant_t < static_cast<long>(index),
			_variant_t(static_cast<double>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, string const& index, double val)
	{
		return impl.add_new(_variant_t < static_cast<const wchar_t*>(index),
			_variant_t(static_cast<double>(val)));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, long index, datetime const& val)
	{
		VARIANT var;
		VariantInit(&var);
		var.vt = static_cast<VARTYPE>(VT_DATE);
		var.date = val;

		_variant_t var_t(var);
		return impl.add_new(_variant_t < static_cast<long>(index),
			var_t);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::add_new(
		implement_type& impl, string const& index, datetime const& val)
	{
		VARIANT var;
		VariantInit(&var);
		var.vt = static_cast<VARTYPE>(VT_DATE);
		var.date = val;

		_variant_t var_t(var);
		return impl.add_new(_variant_t < static_cast<const wchar_t*>(index),
			var_t);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::update(
		implement_type& impl)
	{
		return impl._update();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::update_batch(
		implement_type& impl, msado::affect affectrecords)
	{
		return impl.update_batch(affectrecords);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::cancel_update(
		implement_type& impl)
	{
		impl.cancel_update();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::cancel_batch(
		implement_type& impl, msado::affect affectrecords)
	{
		return impl.cancel_batch(affectrecords);
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_ref_active_connection(
		implement_type& impl, sqlserver_connect<DataBase>& pvar)
	{
		impl.set_ref_active_connection(pvar.get_internel_ptr());
	}
	
	template<typename DataBase>
	inline sqlserver_connect<DataBase> 
		sqlserver_recordset_service<DataBase>::get_ref_active_connection(
		implement_type const& impl) const
	{
		msado::connection conn=impl.get_ref_active_connection();
		return sqlserver_connect<DataBase>(database_, conn);
	}
	
	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_active_connection(
		implement_type& impl, string const& connstr)
	{
		impl.set_active_connection(_variant_t(static_cast<const wchar_t*>(connstr)));
	}
	
	template<typename DataBase>
	inline database::string 
		sqlserver_recordset_service<DataBase>::get_active_connection(
		implement_type const& impl) const
	{

		return impl.get_active_connection();
	}

	//_variant_t get_bookmark(implement_type const& impl) const;
	//void set_bookmark(implement_type& impl, _variant_t const& pvbookmark);

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_cache_size(
		implement_type const& impl) const
	{
		return impl.get_cache_size();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_cache_size(
		implement_type& impl, long pl)
	{
		impl.set_cache_size(static_cast<long>(pl));
	}

	template<typename DataBase>
	inline msado::cursor_type
		sqlserver_recordset_service<DataBase>::get_cursor_type(
		implement_type const& impl) const
	{
		return impl.get_cursor_type();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_cursor_type(
		implement_type& impl, 
		msado::cursor_type plcursortype/* = msado::open_dynamic*/)
	{
		impl.set_cursor_type(plcursortype);
	}

	template<typename DataBase>
	inline msado::lock_type sqlserver_recordset_service<DataBase>::get_lock_type(
		implement_type const& impl) const
	{
		return impl.get_lock_type();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_lock_type(
		implement_type& impl,
		msado::lock_type pllocktype/* = msado::lock_read_only*/)
	{
		impl.set_lock_type(pllocktype);
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_max_records(
		implement_type const& impl) const
	{
		return impl.get_max_records();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_max_records(
		implement_type& impl, long plmaxrecords)
	{
		impl.set_max_records(static_cast<long>(plmaxrecords));
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_ref_source(
		implement_type& impl, sqlserver_command<DataBase>& cmd)
	{
		impl.set_ref_source(cmd.get_internal_ptr());
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_source(
		implement_type& impl, string const& pvsource)
	{
		impl.set_source(pvsource);
	}

	template<typename DataBase>
	inline sqlserver_command<DataBase> 
		sqlserver_recordset_service<DataBase>::get_ref_source(
		implement_type const& impl) const
	{
		msado::command cmd=impl.get_ref_source();
		return sqlserver_command<DataBase>(database_, cmd);
	}

	template<typename DataBase>
	inline database::string sqlserver_recordset_service<DataBase>::get_source(
		implement_type const& impl) const
	{
		return impl.get_source();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::_delete(
		implement_type& impl, msado::affect affectrecords)
	{
		return impl._delete(affectrecords);
	}

	//safearray
	//_variant_t get_rows(implement_type const& impl, long Rows) const;

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::requery(
		implement_type& impl, 
		long options/* = msado::execute_option::option_unspecified*/)
	{
		return impl.requery(static_cast<long>(options));
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::_xresync(
		implement_type& impl, msado::affect affectrecords)
	{
		return impl._xresync(affectrecords);
	}

	template<typename DataBase>
	inline msado::position 
		sqlserver_recordset_service<DataBase>::get_absolute_page(
		implement_type const& impl) const
	{
		return impl.get_absolute_page();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_absolute_page(
		implement_type& impl, msado::position pl/* = msado::pos_bof*/)
	{
		impl.set_absolute_page(pl);
	}

	template<typename DataBase>
	inline msado::edit_mode sqlserver_recordset_service<DataBase>::get_edit_mode(
		implement_type const& impl) const
	{
		return impl.get_edit_mode();
	}


	//_variant_t get_filter(implement_type const& impl) const;

	template<typename DataBase>
	inline string sqlserver_recordset_service<DataBase>::get_condition_filter(
		implement_type const& impl) const
	{
		string strcond;

		_variant_t cond = impl.get_filter();

		if (cond.vt==static_cast<VARTYPE>(VT_BSTR))
		{
			strcond = static_cast<const wchar_t*>(cond);
		}

		return strcond;

	}

	template<typename DataBase>
	inline msado::filter_group
		sqlserver_recordset_service<DataBase>::get_group_filter(
		implement_type const& impl) const
	{
		long lgroup = 0;
		_variant_t group = impl.get_filter();
		if (group.vt==static_cast<VARTYPE>(VT_I4))
		{
			lgroup = group.lVal;
		}

		return static_cast<msado::filter_group>(lgroup);
	}


	//void set_filter(implement_type& impl, const _variant_t & criteria);

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_condition_filter(
		implement_type& impl, string const& condition)
	{
		impl.set_filter(_variant_t(static_cast<const wchar_t*>(condition)));
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_group_filter(
		implement_type& impl, msado::filter_group groupfilter)
	{
		impl.set_filter(_variant_t(static_cast<long>(groupfilter)));
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_page_count(
		implement_type const& impl) const
	{
		return impl.get_page_count();
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_page_size(
		implement_type const& impl) const
	{
		return impl.get_page_size();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_page_size(
		implement_type& impl, long pl)
	{
		impl.set_page_size(static_cast<long>(pl));
	}

	template<typename DataBase>
	inline database::string sqlserver_recordset_service<DataBase>::get_sort(
		implement_type const& impl) const
	{
		return impl.get_sort();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_sort(
		implement_type& impl, string const& criteria)
	{
		impl.set_sort(criteria);
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_status(
		implement_type const& impl) const
	{
		return impl.get_status();
	}

	template<typename DataBase>
	inline long sqlserver_recordset_service<DataBase>::get_state(
		implement_type const& impl) const
	{
		return impl.get_state();
	}

	template<typename DataBase>
	inline sqlserver_recordset<DataBase> 
		sqlserver_recordset_service<DataBase>::_xclone(
		implement_type& impl)
	{
		msado::recordset set=impl._xclone();
		return sqlserver_recordset<DataBase>(database_, set);
	}

	template<typename DataBase>
	inline msado::cursor_location 
		sqlserver_recordset_service<DataBase>::get_cursor_location(
		implement_type const& impl) const
	{
		return impl.get_cursor_location();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_cursor_location(
		implement_type& impl, msado::cursor_location plcursorloc)
	{
		impl.set_cursor_location(plcursorloc);
	}

	template<typename DataBase>
	inline sqlserver_recordset<DataBase> 
		sqlserver_recordset_service<DataBase>::next_recordset(
		implement_type& impl, long * recordsaffected)
	{
		msado::recordset set = impl.next_recordset(recordsaffected);
		return sqlserver_recordset<DataBase>(database_, set);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::supports(
		implement_type& impl, msado::cursor_option cursoroptions)
	{
		return impl.supports(cursoroptions);
	}

	template<typename DataBase>
	inline database::string sqlserver_recordset_service<DataBase>::get_collect(
		implement_type const& impl, long index) const
	{
		string strcol;
		_variant_t col = impl.get_collect(_variant_t(static_cast<long>(index)));
		if (col.vt==static_cast<VARTYPE>(VT_BSTR))
		{
			strcol = col.bstrVal;
		}
		return strcol;
	}

	template<typename DataBase>
	inline database::string sqlserver_recordset_service<DataBase>::get_collect(
		implement_type const& impl, string const& index) const
	{
		string strcol;
		_variant_t col = impl.get_collect(_variant_t(static_cast<const wchar_t*>(index)));
		if (col.vt == static_cast<VARTYPE>(VT_BSTR))
		{
			strcol = col.bstrVal;
		}
		return strcol;
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_collect(
		implement_type& impl, long index, string const& pvar)
	{
		impl.set_collect(_variant_t(static_cast<long>(index)),
			_variant_t(static_cast<const wchar_t*>(pvar)));
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_collect(
		implement_type& impl, string const& index, string const& pvar)
	{
		impl.set_collect(_variant_t(static_cast<wchar_t const*>(index)),
			_variant_t(static_cast<const wchar_t*>(pvar)));
	}

	template<typename DataBase>
	inline msado::marshal_options 
		sqlserver_recordset_service<DataBase>::get_marshal_options(
		implement_type const& impl) const
	{
		return impl.get_marshal_options();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_marshal_options(
		implement_type& impl, msado::marshal_options pemarshal)
	{
		impl.set_marshal_options(pemarshal);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::_xsave(
		implement_type& impl,
		string const& filename, msado::persist_format persistformat)
	{
		return impl._xsave(filename, persistformat);
	}

	template<typename DataBase>
	inline sqlserver_command<DataBase> 
		sqlserver_recordset_service<DataBase>::get_active_command(
		implement_type const& impl) const
	{
		msado::command cmd = impl.get_active_command();
		return sqlserver_command<DataBase>(database_, cmd);
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_stay_in_sync(
		implement_type& impl, bool pbstayinsync)
	{
		impl.set_stay_in_sync(pbstayinsync);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::get_stay_in_sync(
		implement_type const& impl) const
	{
		return impl.get_stay_in_sync();
	}

	template<typename DataBase>
	inline string sqlserver_recordset_service<DataBase>::get_string(
		implement_type const& impl,
		msado::string_format stringformat, long numrows,
		string const& columndelimeter,
		string const& rowdelimeter,
		string const& _nullexpr) const
	{
		return impl.get_string(stringformat, static_cast<long>(numrows),
			columndelimeter, rowdelimeter, _nullexpr);
	}

	template<typename DataBase>
	inline string sqlserver_recordset_service<DataBase>::get_data_member(
		implement_type const& impl) const
	{
		return impl.get_data_member();
	}

	template<typename DataBase>
	inline void sqlserver_recordset_service<DataBase>::set_data_member(
		implement_type& impl, string const& pbstrdatamember)
	{
		impl.set_data_member(pbstrdatamember);
	}

	//msado::_compare compare_bookmarks(implement_type& impl,
	//	_variant_t const& bookmark1, _variant_t const& bookmark2);

	template<typename DataBase>
	inline sqlserver_recordset<DataBase> 
		sqlserver_recordset_service<DataBase>::clone(
		implement_type& impl, msado::lock_type locktype)
	{
		msado::recordset set = impl.clone(locktype);
		return sqlserver_recordset<DataBase>(database_, set);
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::resync(implement_type& impl,
		msado::affect affectrecords, msado::_resync resyncvalues)
	{
		return impl.resync(affectrecords, resyncvalues);
	}

	//template<typename DataBase>
	//inline bool sqlserver_recordset_service<DataBase>::seek(implement_type& impl,
	//	_variant_t const& keyvalues, msado::_seek seekoption/* = seek_first_eq*/)
	//{
	//}

	//void set_index(implement_type& impl, string const& pbstrindex);
	//string get_index(implement_type const& impl) const;
	//bool save(implement_type& impl, _variant_t const& destination,
	//	msado::persist_format persistformat = msado::persist_adtg);

	///////////////////////////////////////////////////////////////////////
	// internal method

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::opened(
		implement_type const& impl) const
	{
		return impl.get_state() != msado::state_closed;
	}


	template<typename DataBase>
	inline _variant_t sqlserver_recordset_service<DataBase>::get_value_impl(
		implement_type const& impl, long index) const
	{
		return impl.get_fields()
			.get_item(_variant_t(static_cast<long>(index), VT_I4))
			.get_value();
	}

	template<typename DataBase>
	inline _variant_t sqlserver_recordset_service<DataBase>::get_value_impl(
		implement_type const& impl, string const& index) const
	{
		return impl.get_fields()
			.get_item(_variant_t(static_cast<const wchar_t*>(index))).get_value();
	}

	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value_impl(
		implement_type& impl, long index, _variant_t const& val)
	{
		if (get_state() == static_cast<long>(msado::edit_none))
		{
			return false;
		}

		impl.get_fields()
			.get_item(_variant_t(static_cast<long>(index), VT_I4))
			.set_value(val);

		return true;
	}
	
	template<typename DataBase>
	inline bool sqlserver_recordset_service<DataBase>::set_value_impl(
		implement_type& impl, string const& index, _variant_t const& val)
	{
		if (get_state() == static_cast<long>(msado::edit_none))
		{
			return false;
		}

		impl.get_fields()
			.get_item(_variant_t(static_cast<const wchar_t*>(index)))
			.set_value(val);

		return true;
	}
}
