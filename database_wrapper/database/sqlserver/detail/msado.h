#pragma once

#include <Windows.h>

#include <atlbase.h>
#include <atlstr.h>
#include <comdef.h>

#include <adoint.h>

#define CLSID_CONNECTION L"ADODB.Connection"
#define CLSID_COMMAND    L"ADODB.Command"
#define CLSID_RECORDSET  L"ADODB.Recordset"
#define CLSID_ERROR      L"ADODB.Error"
#define CLSID_PARAMETER  L"ADODB.Parameter"
#define CLSID_RECORD     L"ADODB.Record"
#define CLSID_STREAM     L"ADODB.Stream"

namespace msado
{
	class Clsid
	{
	public:
		static CLSID ADODB_Connection()
		{
			CLSID clsid;
			CLSIDFromProgID(const_cast<LPWSTR> (CLSID_CONNECTION), &clsid);
			return clsid;
		}

		static CLSID ADODB_Command()
		{
			CLSID clsid;
			CLSIDFromProgID(const_cast<LPWSTR> (CLSID_COMMAND), &clsid);
			return clsid;
		}

		static CLSID ADODB_Recordset()
		{
			CLSID clsid;
			CLSIDFromProgID(const_cast<LPWSTR> (CLSID_RECORDSET), &clsid);
			return clsid;
		}
	};

	//typedef ATL::COleDateTime datetime;
}

_COM_SMARTPTR_TYPEDEF(Property, __uuidof(Property));
_COM_SMARTPTR_TYPEDEF(Properties, __uuidof(Properties));
_COM_SMARTPTR_TYPEDEF(_ADO, __uuidof(_ADO));
_COM_SMARTPTR_TYPEDEF(Error, __uuidof(Error));
_COM_SMARTPTR_TYPEDEF(Errors, __uuidof(Errors));
_COM_SMARTPTR_TYPEDEF(Field20, __uuidof(Field20));
_COM_SMARTPTR_TYPEDEF(Field, __uuidof(Field));
_COM_SMARTPTR_TYPEDEF(Fields15, __uuidof(Fields15));
_COM_SMARTPTR_TYPEDEF(Fields20, __uuidof(Fields20));
_COM_SMARTPTR_TYPEDEF(Fields, __uuidof(Fields));
_COM_SMARTPTR_TYPEDEF(_Parameter, __uuidof(_Parameter));
_COM_SMARTPTR_TYPEDEF(Parameters, __uuidof(Parameters));
_COM_SMARTPTR_TYPEDEF(ConnectionEvents, __uuidof(ConnectionEvents));
_COM_SMARTPTR_TYPEDEF(RecordsetEvents, __uuidof(RecordsetEvents));
_COM_SMARTPTR_TYPEDEF(ADOConnectionConstruction15, __uuidof(ADOConnectionConstruction15));
_COM_SMARTPTR_TYPEDEF(ADOConnectionConstruction, __uuidof(ADOConnectionConstruction));
_COM_SMARTPTR_TYPEDEF(_Stream, __uuidof(_Stream));
_COM_SMARTPTR_TYPEDEF(ADORecordConstruction, __uuidof(ADORecordConstruction));
_COM_SMARTPTR_TYPEDEF(ADOStreamConstruction, __uuidof(ADOStreamConstruction));
_COM_SMARTPTR_TYPEDEF(ADOCommandConstruction, __uuidof(ADOCommandConstruction));
_COM_SMARTPTR_TYPEDEF(ADORecordsetConstruction, __uuidof(ADORecordsetConstruction));
_COM_SMARTPTR_TYPEDEF(Field15, __uuidof(Field15));
_COM_SMARTPTR_TYPEDEF(Command15, __uuidof(Command15));
_COM_SMARTPTR_TYPEDEF(Command25, __uuidof(Command25));
_COM_SMARTPTR_TYPEDEF(_Command, __uuidof(_Command));
_COM_SMARTPTR_TYPEDEF(Connection15, __uuidof(Connection15));
_COM_SMARTPTR_TYPEDEF(_Connection, __uuidof(_Connection));
_COM_SMARTPTR_TYPEDEF(Recordset15, __uuidof(Recordset15));
_COM_SMARTPTR_TYPEDEF(Recordset20, __uuidof(Recordset20));
_COM_SMARTPTR_TYPEDEF(Recordset21, __uuidof(Recordset21));
_COM_SMARTPTR_TYPEDEF(_Recordset, __uuidof(_Recordset));
_COM_SMARTPTR_TYPEDEF(ConnectionEventsVt, __uuidof(ConnectionEventsVt));
_COM_SMARTPTR_TYPEDEF(RecordsetEventsVt, __uuidof(RecordsetEventsVt));
_COM_SMARTPTR_TYPEDEF(_Record, __uuidof(_Record));

namespace msado
{
	enum cursor_type
	{
		open_unspecified = adOpenUnspecified,
		open_forward_only = adOpenForwardOnly,
		open_keyset = adOpenKeyset,
		open_dynamic = adOpenDynamic,
		open_static = adOpenStatic,
	};

	enum cursor_option
	{
		hold_records = adHoldRecords,
		move_pervious = adMovePrevious,
		add_new = adAddNew,
		_delete = adDelete,
		update = adUpdate,
		book_mark = adBookmark,
		approx_position = adApproxPosition,
		update_batch = adUpdateBatch,
		re_sync = adResync,
		notify = adNotify,
		find = adFind,
		seek = adSeek,
		index = adIndex,
	};

	enum lock_type
	{
		lock_unspecified = adLockUnspecified,
		lock_read_only = adLockReadOnly,
		lock_pressimistic = adLockPessimistic,
		lock_optimistic = adLockOptimistic,
		lock_batch_optimistic = adLockBatchOptimistic,
	};

	enum execute_option
	{
		option_unspecified = adOptionUnspecified,
		async_execute = adAsyncExecute,
		async_fetch = adAsyncFetch,
		async_fetch_non_blocking = adAsyncFetchNonBlocking,
		execute_no_records = adExecuteNoRecords,
		execute_stream = adExecuteStream,
		execute_record = adExecuteRecord,
	};

	enum connect_option
	{
		connect_unspecified = adConnectUnspecified,
		async_connect = adAsyncConnect,
	};

	enum object_state
	{
		state_closed = adStateClosed,
		state_open = adStateOpen,
		state_connecting = adStateConnecting,
		state_executing = adStateExecuting,
		state_fetching = adStateFetching,
	};

	enum cursor_location
	{
		use_none = adUseNone,
		use_server = adUseServer,
		use_client = adUseClient,
		use_client_batch = adUseClientBatch,
	};

	enum data_type
	{
		_empty = adEmpty,
		_tiny_int = adTinyInt,
		_small_int = adSmallInt,
		_integer = adInteger,
		_big_int = adBigInt,
		_unsigned_tiny_int = adUnsignedTinyInt,
		_unsigned_small_int = adUnsignedSmallInt,
		_unsigned_int = adUnsignedInt,
		_unsigned_big_int = adUnsignedBigInt,
		_single = adSingle,
		_double = adDouble,
		_currency = adCurrency,
		_decimal = adDecimal,
		_numeric = adNumeric,
		_boolean = adBoolean,
		_error = adError,
		_user_defined = adUserDefined,
		_variant = adVariant,
		_IDispatch = adIDispatch,
		_IUnknown = adIUnknown,
		_guid = adGUID,
		_date = adDate,
		_db_date = adDBDate,
		_db_time = adDBTime,
		_db_timp_stamp = adDBTimeStamp,
		_bstr = adBSTR,
		_char = adChar,
		_var_char = adVarChar,
		_long_var_char = adLongVarChar,
		_wchar = adWChar,
		_var_wchar = adVarWChar,
		_long_var_wchar = adLongVarWChar,
		_binary = adBinary,
		_var_binary = adVarBinary,
		_long_var_binary = adLongVarBinary,
		_chapter = adChapter,
		_file_time = adFileTime,
		_prop_variant = adPropVariant,
		_var_numeric = adVarNumeric,
		_array = adArray,
	};

	enum field_attribute
	{
		fld_unspecified = adFldUnspecified,
		fld_may_defer = adFldMayDefer,
		fld_updatable = adFldUpdatable,
		fld_unknown_updatable = adFldUnknownUpdatable,
		fld_fixed = adFldFixed,
		fld_is_nullable = adFldIsNullable,
		fld_maybe_null = adFldMayBeNull,
		fld_long = adFldLong,
		fld_row_id = adFldRowID,
		fld_row_version = adFldRowVersion,
		fld_cache_defferrd = adFldCacheDeferred,
		fld_is_chapter = adFldIsChapter,
		fld_negative_scale = adFldNegativeScale,
		fld_key_column = adFldKeyColumn,
		fld_is_row_url = adFldIsRowURL,
		fld_is_default_stream = adFldIsDefaultStream,
		fld_is_collection = adFldIsCollection,
	};

	enum affect
	{
		affect_current=adAffectCurrent,
		affect_group=adAffectGroup,
		affect_all=adAffectAll,
		affect_all_chapters=adAffectAllChapters,
	};

	enum position
	{
		pos_unknown=adPosUnknown,
		pos_bof=adPosBOF,
		pos_eof=adPosEOF,
	};

	enum edit_mode
	{
		edit_none = adEditNone,
		edit_in_progress = adEditInProgress,
		edit_add = adEditAdd,
		edit_delete = adEditDelete,
	};

	enum parameter_direction
	{
		param_unknown = adParamUnknown,
		param_input = adParamInput,
		param_output = adParamOutput,
		param_intput_output = adParamInputOutput,
		param_return_value = adParamReturnValue,
	};

	enum marshal_options
	{
		marshal_all = adMarshalAll,
		marshal_modified_only = adMarshalModifiedOnly,
	};

	enum search_direction
	{
		search_forward = adSearchForward,
		search_backward = adSearchBackward,
	};

	enum connect_mode
	{
		mode_unknown = adModeUnknown,
		mode_read = adModeRead,
		mode_write = adModeWrite,
		mode_read_write = adModeReadWrite,
		mode_share_deny_read = adModeShareDenyRead,
		mode_share_deny_write = adModeShareDenyWrite,
		mode_share_exclusive = adModeShareExclusive,
		mode_share_deny_none = adModeShareDenyNone,
		mode_recursive = adModeRecursive,
	};

	enum command_type
	{
		cmd_unspecified = adCmdUnspecified,
		cmd_unknown = adCmdUnknown,
		cmd_text = adCmdText,
		cmd_table = adCmdTable,
		cmd_stored_proc = adCmdStoredProc,
		cmd_file = adCmdFile,
		cmd_table_direct = adCmdTableDirect,
	};

	enum isolation_level
	{
		xact_unspecified = adXactUnspecified,
		xact_chaos = adXactChaos,
		xact_read_uncommitted = adXactReadUncommitted,
		xact_browser = adXactBrowse,
		xact_cursor_stability = adXactCursorStability,
		xact_read_committed = adXactReadCommitted,
		xact_repeatable_read = adXactRepeatableRead,
		xact_serializable = adXactSerializable,
		xact_isolated = adXactIsolated,
	};

	enum filter_group
	{
		filter_none = adFilterNone,
		filter_pending_records = adFilterPendingRecords,
		filter_affected_records = adFilterAffectedRecords,
		filter_fetched_records = adFilterFetchedRecords,
		filter_predicate = adFilterPredicate,
		filter_conflicting_records = adFilterConflictingRecords,
	};

	enum _compare
	{
		compare_less_than = adCompareLessThan,
		compare_equal = adCompareEqual,
		compare_greater_than = adCompareGreaterThan,
		compare_not_equal = adCompareNotEqual,
		compare_not_comparable = adCompareNotComparable,
	};

	enum _resync
	{
		resync_underlying_values = adResyncUnderlyingValues,
		resync_all_values = adResyncAllValues,
	};

	enum _seek
	{
		seek_first_eq = adSeekFirstEQ,
		seek_last_eq = adSeekLastEQ,
		seek_after_eq = adSeekAfterEQ,
		seek_after = adSeekAfter,
		seek_before_eq = adSeekBeforeEQ,
		seek_before = adSeekBefore,
	};

	enum schema
	{
		schema_provider_specific = adSchemaProviderSpecific,
		schema_asserts = adSchemaAsserts,
		schema_catalogs = adSchemaCatalogs,
		schema_character_sets = adSchemaCharacterSets,
		schema_collations = adSchemaCollations,
		schema_columns = adSchemaColumns,
		schema_check_constraints = adSchemaCheckConstraints,
		schema_constraint_column_usage = adSchemaConstraintColumnUsage,
		schema_constraint_table_usage = adSchemaConstraintTableUsage,
		schema_key_column_usage = adSchemaKeyColumnUsage,
		schema_referential_contraints = adSchemaReferentialContraints,
		schema_referential_constraints = adSchemaReferentialConstraints,
		schema_table_constraints = adSchemaTableConstraints,
		schema_columns_domain_usage = adSchemaColumnsDomainUsage,
		schema_indexes = adSchemaIndexes,
		schema_column_privileges = adSchemaColumnPrivileges,
		schema_table_privileges = adSchemaTablePrivileges,
		schema_usage_privileges = adSchemaUsagePrivileges,
		schema_procedures = adSchemaProcedures,
		schema_schemata = adSchemaSchemata,
		schema_sql_languages = adSchemaSQLLanguages,
		schema_statistics = adSchemaStatistics,
		schema_tables = adSchemaTables,
		schema_translations = adSchemaTranslations,
		schema_provider_types = adSchemaProviderTypes,
		schema_views = adSchemaViews,
		schema_view_column_usage = adSchemaViewColumnUsage,
		schema_view_table_usage = adSchemaViewTableUsage,
		schema_procedure_parameters = adSchemaProcedureParameters,
		schema_foreign_keys = adSchemaForeignKeys,
		schema_primary_keys = adSchemaPrimaryKeys,
		schema_procedure_columns = adSchemaProcedureColumns,
		schema_db_info_keywords = adSchemaDBInfoKeywords,
		schema_db_info_literals = adSchemaDBInfoLiterals,
		schema_cubes = adSchemaCubes,
		schema_dimensions = adSchemaDimensions,
		schema_hierarchies = adSchemaHierarchies,
		schema_levels = adSchemaLevels,
		schema_measures = adSchemaMeasures,
		schema_properties = adSchemaProperties,
		schema_members = adSchemaMembers,
		schema_trustees = adSchemaTrustees,
		schema_functions = adSchemaFunctions,
		schema_actions = adSchemaActions,
		schema_commands = adSchemaCommands,
		schema_sets = adSchemaSets,
	};

	enum persist_format
	{
		persist_adtg = adPersistADTG,
		persist_xml = adPersistXML,
	};

	enum string_format
	{
		clip_string=adClipString,
	};

	enum event_status
	{
		status_ok = adStatusOK,
		status_errors_occurred = adStatusErrorsOccurred,
		status_cant_deny = adStatusCantDeny,
		status_cancel = adStatusCancel,
		status_unwanted_event = adStatusUnwantedEvent,
	};

	enum event_reason
	{
		rsn_add_new = adRsnAddNew,
		rsn_delete = adRsnDelete,
		rsn_update = adRsnUpdate,
		rsn_undo_update = adRsnUndoUpdate,
		rsn_undo_add_new = adRsnUndoAddNew,
		rsn_undo_delete = adRsnUndoDelete,
		rsn_requery = adRsnRequery,
		rsn_resynch = adRsnResynch,
		rsn_close = adRsnClose,
		rsn_move = adRsnMove ,
		rsn_first_change = adRsnFirstChange,
		rsn_move_first = adRsnMoveFirst,
		rsn_move_next = adRsnMoveNext,
		rsn_move_previous = adRsnMovePrevious,
		rsn_move_last = adRsnMoveLast,
	};
}

namespace msado
{
	typedef ATL::CString string;

	class connection;
	class command;
	class recordset;

	namespace detail
	{
		string _bstr_t_to_string(_bstr_t const& bstr);
		_bstr_t string_to_bstr_t(string const& str);
	}

	namespace detail
	{
		struct bool_struct
		{
			int unused;
		};
		typedef int bool_struct::*bool_type;

		template<typename T>
		class non_counter : public T
		{
			ULONG __stdcall AddRef();
			ULONG __stdcall Release();
		};

		class cominit
		{
		public:
			cominit();
			~cominit();
		};

		template<typename SmartPointer>
		class object //: cominit
		{
		public:
			typedef typename SmartPointer::Interface interface_type;
			//typedef IDispatch base_interface_type;
			typedef IUnknown base_interface_type;
		public:
			operator bool_type() const;
			//base_interface_type* dispatch() const;
			base_interface_type* unknown() const;
			interface_type* get() const;
			interface_type* get();
			interface_type** get_address_of();
			void reset();
		protected:
			object();
			//object(base_interface_type* ptr);
			object(interface_type* ptr);
			object(object const& rhs);
			virtual ~object();
		protected:
			void copy(object const& rhs);
			void copy(base_interface_type* rhs);
		protected:
			SmartPointer ptr_;
		private:
			bool operator==(object const&);
			bool operator!=(object const&);
		};

		template<typename T, typename SmartPointer, bool bInit = true>
		class cimpl
			: public object<SmartPointer>
		{
		public:
			cimpl();
			explicit cimpl(CLSID const& clsid);
			cimpl(interface_type* ptr);
			cimpl(cimpl const& rhs);
		public:
			T& operator=(T const& rhs);
			T& operator=(interface_type* rhs);
		public:
			non_counter<interface_type>* operator->() const;
		};

		template<typename T, typename SmartPointer>
		class cimpl<T, SmartPointer, false> 
			: public object<SmartPointer>
		{
		public:
			cimpl();
			cimpl(interface_type* ptr);
			cimpl(cimpl const& rhs);
		public:
			T& operator=(T const& rhs);
			T& operator=(interface_type* rhs);
		public:
			non_counter<interface_type>* operator->() const;
		};
	}

	using detail::cominit;
}

namespace msado
{
	namespace detail
	{
		inline string _bstr_t_to_string(_bstr_t const& bstr)
		{
			return static_cast<LPCWSTR>(bstr);
		}

		inline _bstr_t string_to_bstr_t(string const& str)
		{
			return static_cast<LPCWSTR>(str);
		}
	}

	namespace detail
	{
		//////////////////////////////////////////////////////////
		// cominit

		inline cominit::cominit()
		{
			::CoInitialize(NULL);
		}

		inline cominit::~cominit()
		{
			::CoUninitialize();
		}

		//////////////////////////////////////////////////////////
		// object

		template<typename SmartPointer>
		inline object<SmartPointer>::object()
		{

		}

		//template<typename SmartPointer>
		//inline object<SmartPointer>::object(base_interface_type* ptr)
		//	: ptr_(SmartPointer(reinterpret_cast<interface_type*>(ptr)/*,false*/))
		//{

		//}

		template<typename SmartPointer>
		inline object<SmartPointer>::object(interface_type* ptr)
			: ptr_(SmartPointer(ptr))
		{

		}

		template<typename SmartPointer>
		inline object<SmartPointer>::object(object const& rhs)
			: ptr_(rhs.ptr_)
		{

		}

		template<typename SmartPointer>
		inline object<SmartPointer>::~object()
		{

		}

		template<typename SmartPointer>
		inline object<SmartPointer>::operator bool_type() const
		{
			return ptr_ == NULL ? NULL : &bool_struct::unused;
		}

		//template<typename SmartPointer>
		//inline typename object<SmartPointer>::base_interface_type* 
		//    object<SmartPointer>::dispatch() const
		//{
		//	return ptr_.GetInterfacePtr();
		//}

		template<typename SmartPointer>
		inline typename object<SmartPointer>::base_interface_type*
			object<SmartPointer>::unknown() const
		{
			return ptr_.GetInterfacePtr();
		}

		template<typename SmartPointer>
		inline typename object<SmartPointer>::interface_type*
			object<SmartPointer>::get() const
		{
			return ptr_.GetInterfacePtr();
		}

		template<typename SmartPointer>
		inline typename object<SmartPointer>::interface_type*
			object<SmartPointer>::get()
		{
			return ptr_.GetInterfacePtr();
		}

		template<typename SmartPointer>
		inline typename object<SmartPointer>::interface_type** 
			object<SmartPointer>::get_address_of()
		{
			return ptr_.operator&();
		}

		template<typename SmartPointer>
		inline void object<SmartPointer>::reset()
		{
			interface_type*& temp = ptr_.GetInterfacePtr();
			if (temp != NULL)
			{
				temp->Release();
				temp = NULL;
			}
		}

		template<typename SmartPointer>
		inline void object<SmartPointer>::copy(object const& rhs)
		{
			ptr_ = rhs.ptr_;
		}

		template<typename SmartPointer>
		inline void object<SmartPointer>::copy(base_interface_type* rhs)
		{
			ptr_ = reinterpret_cast<interface_type*>(rhs);
		}

		//////////////////////////////////////////////////////////
		// cimpl

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline cimpl<T, SmartPointer, bInit>::cimpl()
			: object<SmartPointer>()
		{
		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline cimpl<T, SmartPointer, bInit>::cimpl(CLSID const& clsid)
			: object<SmartPointer>()
		{
			ptr_.CreateInstance(clsid);
		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline cimpl<T, SmartPointer, bInit>::cimpl(interface_type* ptr)
			: object<SmartPointer>(ptr)
		{

		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline cimpl<T, SmartPointer, bInit>::cimpl(cimpl const& rhs)
			: object<SmartPointer>(rhs)
		{

		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline T& cimpl<T, SmartPointer, bInit>::operator=(T const& rhs)
		{
			copy(rhs);
			return static_cast<T*>(this);
		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline T& cimpl<T, SmartPointer, bInit>::operator=(interface_type* rhs)
		{
			return static_cast<T*>(this);
		}

		template<typename T, typename SmartPointer, bool bInit/*=true*/>
		inline non_counter<typename cimpl<T, SmartPointer, bInit>::interface_type>*
			cimpl<T, SmartPointer, bInit>::operator->() const
		{
			return static_cast<
				non_counter<typename cimpl<T, SmartPointer, bInit>::interface_type>*
			>(get());
		}


		//////////////////////////////////////////////////////////
		// cimpl ÌØ»¯

		template<typename T, typename SmartPointer>
		inline cimpl<T, SmartPointer, false>::cimpl()
			: object<SmartPointer>()
		{
		}

		template<typename T, typename SmartPointer>
		inline cimpl<T, SmartPointer, false>::cimpl(interface_type* ptr)
			: object<SmartPointer>(ptr)
		{

		}

		template<typename T, typename SmartPointer>
		inline cimpl<T, SmartPointer, false>::cimpl(cimpl const& rhs)
			: object<SmartPointer>(rhs)
		{

		}

		template<typename T, typename SmartPointer>
		inline T& cimpl<T, SmartPointer, false>::operator=(T const& rhs)
		{
			copy(rhs);
			return static_cast<T*>(this);
		}

		template<typename T, typename SmartPointer>
		inline T& cimpl<T, SmartPointer, false>::operator=(interface_type* rhs)
		{
			return static_cast<T*>(this);
		}

		template<typename T, typename SmartPointer>
		inline non_counter<typename cimpl<T, SmartPointer, false>::interface_type>*
			cimpl<T, SmartPointer, false>::operator->() const
		{
			return static_cast<
				non_counter<typename cimpl<T, SmartPointer, false>::interface_type>*
			>(get());
		}
	}
}
