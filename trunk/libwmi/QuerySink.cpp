#include "pch.h"
#include "QuerySink.h"


#pragma warning(disable:33005)


ULONG QuerySink::AddRef()
{
    return InterlockedIncrement(&m_lRef);
}


ULONG QuerySink::Release()
{
    LONG lRef = InterlockedDecrement(&m_lRef);
    if (lRef == 0)
        delete this;

    return lRef;
}

HRESULT QuerySink::QueryInterface(REFIID riid, void ** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IWbemObjectSink) {
        *ppv = (IWbemObjectSink *)this;
        AddRef();
        return WBEM_S_NO_ERROR;
    } else {
        return E_NOINTERFACE;
    }
}


HRESULT QuerySink::Indicate(long lObjectCount, IWbemClassObject ** apObjArray)
{
    HRESULT hres = S_OK;

    for (int i = 0; i < lObjectCount; i++) {
        VARIANT varName;
        hres = apObjArray[i]->Get(_bstr_t(L"Name"), 0, &varName, 0, 0);
        if (FAILED(hres)) {
            cout << "Failed to get the data from the query"
                << " Error code = 0x"
                << hex << hres << endl;
            return WBEM_E_FAILED;       // Program has failed.
        }

        printf("Name: %ls\n", V_BSTR(&varName));
    }

    return WBEM_S_NO_ERROR;
}

HRESULT QuerySink::SetStatus(
    /* [in] */ LONG lFlags,
    /* [in] */ HRESULT hResult,
    /* [in] */ BSTR strParam,
    /* [in] */ IWbemClassObject __RPC_FAR * pObjParam
)
{
    if (lFlags == WBEM_STATUS_COMPLETE) {
        printf("Call complete.\n");

        EnterCriticalSection(&threadLock);
        bDone = true;
        LeaveCriticalSection(&threadLock);
    } else if (lFlags == WBEM_STATUS_PROGRESS) {
        printf("Call in progress.\n");
    }

    return WBEM_S_NO_ERROR;
}


bool QuerySink::IsDone()
{
    bool done = true;

    EnterCriticalSection(&threadLock);
    done = bDone;
    LeaveCriticalSection(&threadLock);

    return done;
}    // end of QuerySink.cpp
