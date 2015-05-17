
#include "command.h"
#include "event_command.h"

msado::command_event::command_event(msado::command* pevent)
	: ref_(0)
	, eventp_(pevent)
	, dw_event_(0)
	, pconnPointer(NULL)
	, pconnPointerContainer(NULL)
{
}

msado::command_event::command_event(msado::command_event const& rhs)
	: ref_(rhs.ref_)
	, eventp_(rhs.eventp_)
	, dw_event_(rhs.dw_event_)
	, pconnPointerContainer(rhs.pconnPointerContainer)
	, pconnPointer(rhs.pconnPointer)
{
	init();
}

msado::command_event::~command_event()
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

HRESULT STDMETHODCALLTYPE msado::command_event::QueryInterface(
	REFIID riid,
	void **ppvObject)
{
	if (riid == IID_IUnknown || riid == __uuidof(ConnectionEventsVt))
	{
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	return E_FAIL;
}

ULONG STDMETHODCALLTYPE msado::command_event::AddRef(void)
{
	++ref_;
	return ref_;
}

ULONG STDMETHODCALLTYPE msado::command_event::Release(void)
{
	if (--ref_ == 0)
	{
		delete this;
	}
	return ref_;
}

//////////////////////////////////////////////////////
//

HRESULT STDMETHODCALLTYPE msado::command_event::InfoMessage(
	ADOError* pError, EventStatusEnum* adStatus, _ADOConnection* pConnection)
{
	*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::BeginTransComplete(
	LONG TransactionLevel,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_begin_transed(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::CommitTransComplete(
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_commit_transed(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::RollbackTransComplete(
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_rollback_transed(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::WillExecute(
	BSTR *Source,
	CursorTypeEnum *CursorType,
	LockTypeEnum *LockType,
	long *Options,
	EventStatusEnum *adStatus,
	_ADOCommand *pCommand,
	_ADORecordset *pRecordset,
	_ADOConnection *pConnection)
{
	eventp_->on_executing(*Source,
		static_cast<cursor_type>(*CursorType),
		static_cast<lock_type>(*LockType),
		*Options,
		pCommand,
		pRecordset,
		pConnection);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::ExecuteComplete(
	LONG RecordsAffected,
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADOCommand *pCommand,
	_ADORecordset *pRecordset,
	_ADOConnection *pConnection)
{
	eventp_->on_executed(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::WillConnect(
	BSTR *ConnectionString,
	BSTR *UserID,
	BSTR *Password,
	long *Options,
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_connecting(*ConnectionString, *UserID, *Password, *Options);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE msado::command_event::ConnectComplete(
	ADOError *pError,
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_connected(pError);
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

void msado::command_event::init()
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
					__uuidof(ConnectionEvents), &pconnPointer);
			}
		}
	}
}

HRESULT STDMETHODCALLTYPE msado::command_event::Disconnect(
	EventStatusEnum *adStatus,
	_ADOConnection *pConnection)
{
	eventp_->on_disconnected();
	//*adStatus = adStatusUnwantedEvent;
	return S_OK;
}

bool msado::command_event::enable_event(bool enable/* = true*/)
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
