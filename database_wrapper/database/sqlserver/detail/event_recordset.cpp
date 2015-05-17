
#include "recordset.h"
#include "event_recordset.h"


msado::recordset_event::recordset_event(msado::recordset* pevent)
	: ref_(1)
	, eventp_(pevent)
	, dw_event_(0)
	, pconnPointer(NULL)
	, pconnPointerContainer(NULL)
{
	
}

msado::recordset_event::recordset_event(msado::recordset_event const& rhs)
	: ref_(rhs.ref_)
	, eventp_(rhs.eventp_)
	, dw_event_(rhs.dw_event_)
	, pconnPointerContainer(rhs.pconnPointerContainer)
	, pconnPointer(rhs.pconnPointer)
{
	init();
}

msado::recordset_event::~recordset_event()
{
	if (pconnPointer)
	{
		pconnPointer->Release();
		pconnPointer = NULL;
	}
	if (pconnPointerContainer)
	{
		pconnPointerContainer->Release();
		pconnPointer = NULL;
	}
}
//////////////////////////////////////////////////////
//

HRESULT STDMETHODCALLTYPE msado::recordset_event::QueryInterface(
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

ULONG STDMETHODCALLTYPE msado::recordset_event::AddRef(void)
{
	++ref_;
	return ref_;
}

ULONG STDMETHODCALLTYPE msado::recordset_event::Release(void)
{
	if (--ref_ == 0)
	{
		delete this;
	}
	return ref_;
}

//////////////////////////////////////////////////////
//

HRESULT STDMETHODCALLTYPE msado::recordset_event::WillChangeField(
	LONG cFields,
	VARIANT Fields,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	//event_.on_field_changing(cFields, _variant_t(Fields, false));
	*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::FieldChangeComplete(
	LONG cFields,
	VARIANT Fields,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_field_changed(cFields, _variant_t(Fields, false), pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::WillChangeRecord(
	EventReasonEnum adReason,
	LONG cRecords,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_record_changing(static_cast<event_reason>(adReason), cRecords);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::RecordChangeComplete(
	EventReasonEnum adReason,
	LONG cRecords,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_record_changed(static_cast<event_reason>(adReason), cRecords, pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::WillChangeRecordset(
	EventReasonEnum adReason,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_recordset_changing(static_cast<event_reason>(adReason));
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::RecordsetChangeComplete(
	EventReasonEnum adReason,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_recordset_changed(static_cast<event_reason>(adReason), pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::WillMove(
	EventReasonEnum adReason,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_moving(static_cast<event_reason>(adReason));
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::MoveComplete(
	EventReasonEnum adReason,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_moved(static_cast<event_reason>(adReason), pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::EndOfRecordset(
	VARIANT_BOOL *fMoreData,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_recordset_end(*fMoreData ? true : false);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::FetchProgress(
	long Progress,
	long MaxProgress,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_fetching(Progress, MaxProgress);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::recordset_event::FetchComplete(
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADORecordset *pRecordset)
{
	eventp_->on_fetched(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

void msado::recordset_event::init()
{
	if (eventp_->get())
	{
		HRESULT hr;
		if (!pconnPointerContainer)
		{
			hr = eventp_->get()->QueryInterface(
				IID_IConnectionPointContainer, (void**)&pconnPointerContainer);
			if (!pconnPointer && SUCCEEDED(hr))
			{
				hr = pconnPointerContainer->FindConnectionPoint(
					__uuidof(RecordsetEvents), &pconnPointer);
			}
		}		
	}
}

bool msado::recordset_event::enable_event(bool enable/* = true*/)
{
	if (enable)
	{
		if (!pconnPointer)return false;

		IUnknown* pEvent;
		HRESULT hr = this->QueryInterface(IID_IUnknown, (void**)&pEvent);
		if (FAILED(hr))return false;

		hr = pconnPointer->Advise(pEvent, &dw_event_);
		if (FAILED(hr))return false;
		pEvent->Release();

		return hr == S_OK;
	}
	else
	{
		if (dw_event_ == 0)return false;
		if (!pconnPointer)return false;

		HRESULT hr = pconnPointer->Unadvise(dw_event_);
		return hr == S_OK;
	}
}
