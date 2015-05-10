#pragma once

#include "msado.h"

namespace msado
{
	template<typename Event>
	class recordset_event
		: public RecordsetEventsVt
	{
	public:
		recordset_event();
	public:
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(
			/* [in] */ REFIID riid,
			/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
		virtual ULONG STDMETHODCALLTYPE AddRef(void);
		virtual ULONG STDMETHODCALLTYPE Release(void);
	public:
		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeField(
			/* [in] */ LONG cFields,
			/* [in] */ VARIANT Fields,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);
		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FieldChangeComplete(
			/* [in] */ LONG cFields,
			/* [in] */ VARIANT Fields,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeRecord(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ LONG cRecords,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecordChangeComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ LONG cRecords,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeRecordset(
			/* [in] */ EventReasonEnum adReason,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecordsetChangeComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillMove(
			/* [in] */ EventReasonEnum adReason,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MoveComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE EndOfRecordset(
			/* [out][in] */ __RPC__inout VARIANT_BOOL *fMoreData,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FetchProgress(
			/* [in] */ long Progress,
			/* [in] */ long MaxProgress,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FetchComplete(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);
	private:
		ULONG ref_;
		Event event_;
	};
}

namespace msado
{
	template<typename Event>
	inline recordset_event<Event>::recordset_event()
		: ref_(0)
	{

	}

	//////////////////////////////////////////////////////
	//

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::QueryInterface(
		REFIID riid,
		void **ppvObject)
	{
		if (riid == IID_IUnknown || riid == __uuidof(RecordsetEventsVt))
		{
			*ppvObject = this;
			AddRef();
			return S_OK;
		}

		return E_FAIL;
	}

	template<typename Event>
	inline ULONG STDMETHODCALLTYPE recordset_event<Event>::AddRef(void)
	{
		++ref_;
		return ref_;
	}

	template<typename Event>
	inline ULONG STDMETHODCALLTYPE recordset_event<Event>::Release(void)
	{
		if (--ref_ == 0)
		{
			delete this;
		}
		return ref_;
	}

	//////////////////////////////////////////////////////
	//

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::WillChangeField(
		LONG cFields,
		VARIANT Fields,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_field_changing(cFields, _variant_t(Fields, false));
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::FieldChangeComplete(
		LONG cFields,
		VARIANT Fields,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_field_changed(cFields, _variant_t(Fields, false), pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::WillChangeRecord(
		EventReasonEnum adReason,
		LONG cRecords,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_record_changing(static_cast<event_reason>(adReason),cRecords);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::RecordChangeComplete(
		EventReasonEnum adReason,
		LONG cRecords,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_record_changed(static_cast<event_reason>(adReason), cRecords, pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::WillChangeRecordset(
		EventReasonEnum adReason,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_recordset_changing(static_cast<event_reason>(adReason));
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::RecordsetChangeComplete(
		EventReasonEnum adReason,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_recordset_changed(static_cast<event_reason>(adReason), pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::WillMove(
		EventReasonEnum adReason,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_moving(static_cast<event_reason>(adReason));
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::MoveComplete(
		EventReasonEnum adReason,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_moved(static_cast<event_reason>(adReason), pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::EndOfRecordset(
		VARIANT_BOOL *fMoreData,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_recordset_end(*fMoreData ? true : false);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::FetchProgress(
		long Progress,
		long MaxProgress,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_fetching(Progress, MaxProgress);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE recordset_event<Event>::FetchComplete(
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADORecordset *pRecordset)
	{
		event_.on_fetched(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}
}
