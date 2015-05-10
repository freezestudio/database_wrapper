#pragma once

#include "msado.h"
#include "properties.h"
#include "fields.h"
#include "error.h"
#include "command.h"
#include "connection.h"

namespace msado
{
	class recordset : public detail::cimpl<recordset, _RecordsetPtr>
	{		
	public:
		recordset();
		recordset(recordset const&);
		recordset(interface_type*);
		explicit recordset(CLSID const& clsid);
		~recordset();
	public:
		properties get_properties() const;
		position get_absolute_position() const;
		void set_absolute_position(position pl);
		void set_ref_active_connection(msado::connection& pvar);
		connection get_ref_active_connection() const;
		void set_active_connection(_variant_t const& pvar);
		string get_active_connection() const;
		bool get_BOF() const;
		_variant_t get_bookmark() const;
		void set_bookmark(_variant_t const& pvBookmark);
		long get_cache_size() const;
		void set_cache_size(long pl);
		cursor_type get_cursor_type() const;
		void set_cursor_type(cursor_type plCursorType = open_dynamic);
		bool get_EOF() const;
		fields get_fields() const;
		lock_type get_lock_type() const;
		void set_lock_type(lock_type plLockType = lock_read_only);
		long get_max_records() const;
		void set_max_records(long plMaxRecords);
		long get_record_count() const;
		void set_ref_source(command const& pvSource);
		command get_ref_source() const;
		void set_source(string const& pvSource);
		string get_source() const;
		bool add_new(
			_variant_t const& vtFieldList = vtMissing,
			_variant_t const& vtValues = vtMissing);
		bool cancel_update();
		bool close();
		bool _delete(affect AffectRecords);
		_variant_t get_rows(
			long Rows,
			const _variant_t & Start = vtMissing,
			const _variant_t & Fields = vtMissing);
		bool move(
			long NumRecords,
			const _variant_t & Start = vtMissing);
		bool move_next();
		bool move_previous();
		bool move_first();
		bool move_last();
		bool open(
			const _variant_t & Source,
			const _variant_t & ActiveConnection,
			cursor_type CursorType = open_dynamic,
			lock_type LockType = lock_optimistic,
			long Options = /*command_type::*/cmd_text);
		bool requery(long Options = /*execute_option::*/option_unspecified);
		bool _xresync(affect AffectRecords);
		bool _update(
			const _variant_t & Fields = vtMissing,
			const _variant_t & Values = vtMissing);
		position get_absolute_page() const;
		void set_absolute_page(position pl = pos_bof);
		edit_mode get_edit_mode() const;
		_variant_t get_filter() const;
		void set_filter(const _variant_t & Criteria);
		long get_page_count() const;
		long get_page_size() const;
		void set_page_size(long pl);
		string get_sort() const;
		void set_sort(string const& Criteria);
		long get_status() const;
		long get_state() const;
		recordset _xclone();
		bool update_batch(affect AffectRecords);
		bool cancel_batch(affect AffectRecords);
		cursor_location get_cursor_location() const;
		void set_cursor_location(cursor_location plCursorLoc);
		recordset next_recordset(long * RecordsAffected);
		bool supports(cursor_option CursorOptions);
		_variant_t get_collect(const _variant_t & Index) const;
		void set_collect(
			const _variant_t & Index,
			const _variant_t & pvar);
		marshal_options get_marshal_options() const;
		void set_marshal_options(marshal_options peMarshal);
		bool find(
			string const& Criteria,
			long SkipRecords,
			search_direction SearchDirection = search_forward,
			const _variant_t & Start = vtMissing);
		bool cancel();
		IUnknownPtr get_data_source() const;
		void set_ref_data_source(IUnknown * ppunkDataSource);
		bool _xsave(
			string const& FileName,
		    persist_format PersistFormat);
		command get_active_command() const;
		void set_stay_in_sync(bool pbStayInSync);
		bool get_stay_in_sync() const;
		string get_string(
		    string_format StringFormat, //=clip_string
			long NumRows,
			string const& ColumnDelimeter,
			string const& RowDelimeter,
			string const& NullExpr) const;
		string get_data_member() const;
		void set_data_member(string const& pbstrDataMember);
		_compare compare_bookmarks(
			const _variant_t & Bookmark1,
			const _variant_t & Bookmark2);
		recordset clone(lock_type LockType);
		bool resync(
		    affect AffectRecords,
		    _resync ResyncValues);
		//搜索 Recordset 的索引，快速定位与指定值相匹配的行，并将当前行更改为该行。
		bool seek(
			const _variant_t & KeyValues,
			_seek SeekOption = seek_first_eq);
		void set_index(string const& pbstrIndex);
		string get_index() const;
		bool save(
			const _variant_t & Destination,
			persist_format PersistFormat=persist_adtg);

		//事件
	public:
		void on_field_changing(long cFields, _variant_t const& Fields);
		void on_field_changed(long cFields, _variant_t const& Fields,error const& e);
		void on_record_changing(event_reason adReason, long cRecords);
		void on_record_changed(event_reason adReason, long cRecords, error const& e);
		void on_recordset_changing(event_reason adReason);
		void on_recordset_changed(event_reason adReason,error const& e);
		void on_moving(event_reason adReason);
		void on_moved(event_reason adReason, error const& e);
		void on_recordset_end(bool fMoreData);
		void on_fetching(long Progress, long MaxProgress);
		void on_fetched(error const& e);
	private:
		bool enable_event(bool enable = true);
	private:
		IUnknownPtr eventptr_;
		DWORD dw_event_;
	};
}
