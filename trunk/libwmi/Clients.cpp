#include "pch.h"
#include "Clients.h"

#include "eventsink.h"
#include "querysink.h"


#pragma warning(disable:26495)
#pragma warning(disable:33005)
#pragma warning(disable:4267)
#pragma warning(disable:4244)
#pragma warning(disable:6385)
#pragma warning(disable:6386)
#pragma warning(disable:6011)
#pragma warning(disable:6031)
#pragma warning(disable:6387)


//////////////////////////////////////////////////////////////////////////////////////////////////
//这里是几个MSDN上的例子代码。


EXTERN_C
__declspec(dllexport)
int WINAPI CallingProviderMethod(int iArgCnt, char ** argv)
/*
Example: Calling a Provider Method
2018/05/31

You can use the procedure and code examples in this topic to create a complete WMI client application that performs COM initialization,
connects to WMI on the local computer, calls a provider method, and then cleans up.
The Win32_Process::Create method is used to start Notepad.exe in a new process.

The following procedure is used to execute the WMI application.
Steps 1 through 5 contain all of the steps needed to set up and connect to WMI,
and 6 is where the provider method is called.

To call a provider method

Initialize COM parameters with a call to CoInitializeEx.

For more information, see Initializing COM for a WMI Application.

Initialize COM process security by calling CoInitializeSecurity.

For more information, see Setting the Default Process Security Level Using C++.

Obtain the initial locator to WMI by calling CoCreateInstance.

For more information, see Creating a Connection to a WMI Namespace.

Obtain a pointer to IWbemServices for the root\cimv2 namespace on the local computer by calling IWbemLocator::ConnectServer.
To connect to a remote computer, see Example: Getting WMI Data from a Remote Computer.

For more information, see Creating a Connection to a WMI Namespace.

Set IWbemServices proxy security so the WMI service can impersonate the client by calling CoSetProxyBlanket.

For more information, see Setting the Security Levels on a WMI Connection.

Use the IWbemServices pointer to make requests to WMI.
This example uses IWbemServices::ExecMethod to call the provider method Win32_Process::Create.

For more information about making requests to WMI,
see Manipulating Class and Instance Information and Calling a Method.

If the provider method has any in-parameters or out-parameters,
then values of the parameters must be given to IWbemClassObject pointers.
For in-parameters, you must spawn an instance of the in-parameter definitions,
and then set the values of these new instances.
The Win32_Process::Create method requires a value for the CommandLine in-parameter to execute properly.

The following code example creates an IWbemClassObject pointer,
spawns a new instance of the Win32_Process::Create in-parameter definitions,
and then sets the value of the CommandLine in-parameter to Notepad.exe.

// Set up to call the Win32_Process::Create method
BSTR MethodName = SysAllocString(L"Create");
BSTR ClassName = SysAllocString(L"Win32_Process");

IWbemClassObject* pClass = NULL;
hres = pSvc->GetObject(ClassName, 0, NULL, &pClass, NULL);

IWbemClassObject* pInParamsDefinition = NULL;
hres = pClass->GetMethod(MethodName, 0,
    &pInParamsDefinition, NULL);

IWbemClassObject* pClassInstance = NULL;
hres = pInParamsDefinition->SpawnInstance(0, &pClassInstance);

// Create the values for the in-parameters
VARIANT varCommand;
varCommand.vt = VT_BSTR;
varCommand.bstrVal = _bstr_t(L"notepad.exe");

// Store the value for the in-parameters
hres = pClassInstance->Put(L"CommandLine", 0,
    &varCommand, 0);
wprintf(L"The command is: %s\n", V_BSTR(&varCommand));
The following code example shows how the Win32_Process::Create method out-parameters are given to an IWbemClassObject pointer.
The out-parameter value is obtained with the IWbemClassObject::Get method and stored in a VARIANT variable so that it can be displayed to the user.

// Execute Method
IWbemClassObject* pOutParams = NULL;
hres = pSvc->ExecMethod(ClassName, MethodName, 0,
    NULL, pClassInstance, &pOutParams, NULL);

VARIANT varReturnValue;
hres = pOutParams->Get(_bstr_t(L"ReturnValue"), 0,
    &varReturnValue, NULL, 0);
The following code example shows how to call a provider method using WMI.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example--calling-a-provider-method
*/
{
    HRESULT hres;

    // Step 1: Initialize COM. 
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
        return 1;// Program has failed.
    }

    // Step 2: Set general COM security levels 
    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM negotiates service
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;// Program has failed.
    }

    // Step 3: Obtain the initial locator to WMI
    IWbemLocator * pLoc = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object. " << "Err code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices * pSvc = NULL;

    // Connect to the local root\cimv2 namespace and obtain pointer pSvc to make IWbemServices calls.
    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Step 5: Set security levels for the proxy 
    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 6: Use the IWbemServices pointer to make requests of WMI 

    // set up to call the Win32_Process::Create method
    BSTR MethodName = SysAllocString(L"Create");
    BSTR ClassName = SysAllocString(L"Win32_Process");

    IWbemClassObject * pClass = NULL;
    hres = pSvc->GetObject(ClassName, 0, NULL, &pClass, NULL);

    IWbemClassObject * pInParamsDefinition = NULL;
    hres = pClass->GetMethod(MethodName, 0, &pInParamsDefinition, NULL);

    IWbemClassObject * pClassInstance = NULL;
    hres = pInParamsDefinition->SpawnInstance(0, &pClassInstance);

    // Create the values for the in parameters
    VARIANT varCommand;
    varCommand.vt = VT_BSTR;
    varCommand.bstrVal = _bstr_t(L"notepad.exe");

    // Store the value for the in parameters
    hres = pClassInstance->Put(L"CommandLine", 0, &varCommand, 0);
    wprintf(L"The command is: %s\n", V_BSTR(&varCommand));

    // Execute Method
    IWbemClassObject * pOutParams = NULL;
    hres = pSvc->ExecMethod(ClassName, MethodName, 0, NULL, pClassInstance, &pOutParams, NULL);
    if (FAILED(hres)) {
        cout << "Could not execute method. Error code = 0x" << hex << hres << endl;
        VariantClear(&varCommand);
        SysFreeString(ClassName);
        SysFreeString(MethodName);
        pClass->Release();
        pClassInstance->Release();
        pInParamsDefinition->Release();
        pOutParams->Release();
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // To see what the method returned, use the following code. 
    // The return value will be in &varReturnValue
    VARIANT varReturnValue;
    hres = pOutParams->Get(_bstr_t(L"ReturnValue"), 0, &varReturnValue, NULL, 0);

    // Clean up
    VariantClear(&varCommand);
    VariantClear(&varReturnValue);
    SysFreeString(ClassName);
    SysFreeString(MethodName);
    pClass->Release();
    pClassInstance->Release();
    pInParamsDefinition->Release();
    pOutParams->Release();
    pLoc->Release();
    pSvc->Release();
    CoUninitialize();
    return 0;
}


int ReceivingEventNotifications(int iArgCnt, char ** argv)
/*
Example: Receiving Event Notifications Through WMI
2018/05/31

You can use the procedure and code example in this topic to complete WMI client application that performs COM initialization,
connects to WMI on the local computer, receives an event notification, and then cleans up.
The example notifies the user of an event when a new process is created.
The events are received asynchronously.

The following procedure is used to execute the WMI application.
Steps 1 through 5 contain all of the steps required to set up and connect to WMI,
and step 6 is where the event notifications are received.

To receive a event notification through WMI

Initialize COM parameters with a call to CoInitializeEx.

Fore more information, see Initializing COM for a WMI Application.

Initialize COM process security by calling CoInitializeSecurity.

For more information, see Setting the Default Process Security Level Using C++.

Obtain the initial locator to WMI by calling CoCreateInstance.

For more information, see Creating a Connection to a WMI Namespace.

Obtain a pointer to IWbemServices for the root\cimv2 namespace on the local computer by calling IWbemLocator::ConnectServer.
To connect to a remote computer, see Example: Getting WMI Data from a Remote Computer.

For more information, see Creating a Connection to a WMI Namespace.

Set IWbemServices proxy security so the WMI service can impersonate the client by calling CoSetProxyBlanket.

For more information, see Setting the Security Levels on a WMI Connection.

Use the IWbemServices pointer to make requests to WMI.
This example uses the IWbemServices::ExecNotificationQueryAsync method to receive asynchronous events.
When you receive asynchronous events, you must provide an implementation of IWbemObjectSink.
This example provides the implementation in the EventSink class.
The implementation code and header file code for this class are provided below the main example.
The IWbemServices::ExecNotificationQueryAsync method calls the EventSink::Indicate method whenever an event is received.
For this example the EventSink::Indicate method is called whenever a process is created.
To test this example, run the code and start a process such as Notepad.exe.
This triggers an event notification.

For more information about making requests of WMI,
see Manipulating Class and Instance Information and Calling a Method.

The following example code receives event notifications through WMI.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example--receiving-event-notifications-through-wmi-
*/
{
    HRESULT hres;

    // Step 1: Initialize COM. 
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Step 2: Set general COM security levels 
    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM negotiates service
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                      // Program has failed.
    }

    // Step 3: Obtain the initial locator to WMI 
    IWbemLocator * pLoc = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object. " << "Err code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices * pSvc = NULL;

    // Connect to the local root\cimv2 namespace and obtain pointer pSvc to make IWbemServices calls.
    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Step 5: Set security levels on the proxy 
    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 6: Receive event notifications 

    // Use an unsecured apartment for security
    IUnsecuredApartment * pUnsecApp = NULL;

    hres = CoCreateInstance(CLSID_UnsecuredApartment, NULL,
                            CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment,
                            (void **)&pUnsecApp);

    EventSink * pSink = new EventSink;
    pSink->AddRef();

    IUnknown * pStubUnk = NULL;
    pUnsecApp->CreateObjectStub(pSink, &pStubUnk);

    IWbemObjectSink * pStubSink = NULL;
    pStubUnk->QueryInterface(IID_IWbemObjectSink, (void **)&pStubSink);

    // The ExecNotificationQueryAsync method will call
    // The EventQuery::Indicate method when an event occurs
    hres = pSvc->ExecNotificationQueryAsync(
        _bstr_t("WQL"),
        _bstr_t("SELECT * "
                "FROM __InstanceCreationEvent WITHIN 1 "
                "WHERE TargetInstance ISA 'Win32_Process'"),
        WBEM_FLAG_SEND_STATUS,
        NULL,
        pStubSink);

    // Check for errors.
    if (FAILED(hres)) {
        printf("ExecNotificationQueryAsync failed with = 0x%X\n", hres);
        pSvc->Release();
        pLoc->Release();
        pUnsecApp->Release();
        pStubUnk->Release();
        pSink->Release();
        pStubSink->Release();
        CoUninitialize();
        return 1;
    }

    // Wait for the event
    Sleep(10000);

    hres = pSvc->CancelAsyncCall(pStubSink);

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pUnsecApp->Release();
    pStubUnk->Release();
    pSink->Release();
    pStubSink->Release();
    CoUninitialize();

    return 0;   // Program successfully completed.
}


int GettingDatafromLocalComputerAsynchronously(int argc, char ** argv)
/*
Example: Getting WMI Data from the Local Computer Asynchronously
2018/05/31

You can use the procedure and code examples in this topic to create a complete WMI client application that performs COM initialization,
connects to WMI on the local computer, gets data asynchronously, and then cleans up.
This example gets the name of the operating system on the local computer and displays it.

The following procedure is used to execute the WMI application.
Steps 1 through 5 contain all of the steps required to set up and connect to WMI,
and steps 6 and seven are where the operating system name is retrieved asynchronously.

To get WMI data from the local computer asynchronously

Initialize COM parameters with a call to CoInitializeEx.

For more information, see Initializing COM for a WMI Application.

Initialize COM process security by calling CoInitializeSecurity.

For more information, see Setting the Default Process Security Level Using C++.

Obtain the initial locator to WMI by calling CoCreateInstance.

For more information, see Creating a Connection to a WMI Namespace.

Obtain a pointer to IWbemServices for the root\cimv2 namespace on the local computer by calling IWbemLocator::ConnectServer.
For more information about how to connect to a remote computer,
see Example: Getting WMI Data from a Remote Computer.

For more information, see Creating a Connection to a WMI Namespace.

Set IWbemServices proxy security so the WMI service can impersonate the client by calling CoSetProxyBlanket.

For more information, see Setting the Security Levels on a WMI Connection.

Use the IWbemServices pointer to make requests to WMI.
This example uses the IWbemServices::ExecQueryAsync method to receive the data asynchronously.
Whenever you receive data asynchronously, you must provide an implementation of IWbemObjectSink.
This example provides the implementation in the QuerySink class.
The implementation code and header file code for this class are provided following the main example.
The IWbemServices::ExecQueryAsync method calls the QuerySink::Indicate method whenever the data is received.

For more information about how to create a WMI request,
see Manipulating Class and Instance Information and Calling a Method.

Wait for the data to be retrieved asynchronously.
Use the IWbemServices::CancelAsyncCall method to manually stop the asynchronous call.

The following code example gets WMI data from the local computer asynchronously.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example--getting-wmi-data-from-the-local-computer-asynchronously
*/
{
    HRESULT hres;

    // Step 1: Initialize COM. 
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Step 2: Set general COM security levels 
    hres = CoInitializeSecurity(NULL,
                                -1,                          // COM authentication
                                NULL,                        // Authentication services
                                NULL,                        // Reserved
                                RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
                                RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
                                NULL,                        // Authentication info
                                EOAC_NONE,                   // Additional capabilities 
                                NULL);                       // Reserved
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                    // Program has failed.
    }

    // Step 3: Obtain the initial locator to WMI 
    IWbemLocator * pLoc = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object." << " Err code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices * pSvc = NULL;

    // Connect to the local root\cimv2 namespace and obtain pointer pSvc to make IWbemServices calls.
    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Step 5: Set security levels on the proxy 
    hres = CoSetProxyBlanket(pSvc,                        // Indicates the proxy to set
                             RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
                             RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
                             NULL,                        // Server principal name 
                             RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
                             RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
                             NULL,                        // client identity
                             EOAC_NONE);                  // proxy capabilities 
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 6: Use the IWbemServices pointer to make requests of WMI 

    // For example, get the name of the operating system.
    // The IWbemService::ExecQueryAsync method will call
    // the QuerySink::Indicate method when it receives a result
    // and the QuerySink::Indicate method will display the OS name
    QuerySink * pResponseSink = new QuerySink();
    pResponseSink->AddRef();
    hres = pSvc->ExecQueryAsync(bstr_t("WQL"),
                                bstr_t("SELECT * FROM Win32_OperatingSystem"),
                                WBEM_FLAG_BIDIRECTIONAL,
                                NULL,
                                pResponseSink);
    if (FAILED(hres)) {
        cout << "Query for operating system name failed." << " Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        pResponseSink->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 7: Wait to get the data from the query in step 6 
    Sleep(500);
    pSvc->CancelAsyncCall(pResponseSink);

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return 0;   // Program successfully completed.
}


int __cdecl GettingDatafromRemoteComputer(int argc, char ** argv)
/*
Example: Getting WMI Data from a Remote Computer
2018/05/31

The following code example shows how to get WMI data semisynchronously from a remote computer.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example--getting-wmi-data-from-a-remote-computer
*/
{
    HRESULT hres;

    // Step 1: Initialize COM. 
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Step 2: Set general COM security levels
    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IDENTIFY,    // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                    // Program has failed.
    }

    // Step 3: Obtain the initial locator to WMI 
    IWbemLocator * pLoc = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object." << " Err code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices * pSvc = NULL;

    // Get the user name and password for the remote computer
    CREDUI_INFO cui;
    bool useToken = false;
    bool useNTLM = true;
    wchar_t pszName[CREDUI_MAX_USERNAME_LENGTH + 1] = {0};
    wchar_t pszPwd[CREDUI_MAX_PASSWORD_LENGTH + 1] = {0};
    wchar_t pszDomain[CREDUI_MAX_USERNAME_LENGTH + 1];
    wchar_t pszUserName[CREDUI_MAX_USERNAME_LENGTH + 1];
    wchar_t pszAuthority[CREDUI_MAX_USERNAME_LENGTH + 1];
    BOOL fSave;
    DWORD dwErr;

    memset(&cui, 0, sizeof(CREDUI_INFO));
    cui.cbSize = sizeof(CREDUI_INFO);
    cui.hwndParent = NULL;
    // Ensure that MessageText and CaptionText identify
    // what credentials to use and which application requires them.
    cui.pszMessageText = TEXT("Press cancel to use process token");
    cui.pszCaptionText = TEXT("Enter Account Information");
    cui.hbmBanner = NULL;
    fSave = FALSE;

    dwErr = CredUIPromptForCredentials(
        &cui,                             // CREDUI_INFO structure
        TEXT(""),                         // Target for credentials
        NULL,                             // Reserved
        0,                                // Reason
        pszName,                          // User name
        CREDUI_MAX_USERNAME_LENGTH + 1,     // Max number for user name
        pszPwd,                           // Password
        CREDUI_MAX_PASSWORD_LENGTH + 1,     // Max number for password
        &fSave,                           // State of save check box
        CREDUI_FLAGS_GENERIC_CREDENTIALS |// flags
        CREDUI_FLAGS_ALWAYS_SHOW_UI |
        CREDUI_FLAGS_DO_NOT_PERSIST);
    if (dwErr == ERROR_CANCELLED) {
        useToken = true;
    } else if (dwErr) {
        cout << "Did not get credentials " << dwErr << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // change the computerName strings below to the full computer name of the remote computer
    if (!useNTLM) {
        StringCchPrintf(pszAuthority, CREDUI_MAX_USERNAME_LENGTH + 1, L"kERBEROS:%s", L"COMPUTERNAME");
    }

    // Connect to the remote root\cimv2 namespace and obtain pointer pSvc to make IWbemServices calls.
    hres = pLoc->ConnectServer(_bstr_t(L"\\\\COMPUTERNAME\\root\\cimv2"),
                               _bstr_t(useToken ? NULL : pszName),    // User name
                               _bstr_t(useToken ? NULL : pszPwd),     // User password
                               NULL,                              // Locale             
                               NULL,                              // Security flags
                               _bstr_t(useNTLM ? NULL : pszAuthority),// Authority        
                               NULL,                              // Context object 
                               &pSvc                              // IWbemServices proxy
    );
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // step 5: Create COAUTHIDENTITY that can be used for setting security on proxy
    COAUTHIDENTITY * userAcct = NULL;
    COAUTHIDENTITY authIdent;

    if (!useToken) {
        memset(&authIdent, 0, sizeof(COAUTHIDENTITY));
        authIdent.PasswordLength = wcslen(pszPwd);
        authIdent.Password = (USHORT *)pszPwd;

        LPWSTR slash = wcschr(pszName, L'\\');
        if (slash == NULL) {
            cout << "Could not create Auth identity. No domain specified\n";
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            return 1;               // Program has failed.
        }

        StringCchCopy(pszUserName, CREDUI_MAX_USERNAME_LENGTH + 1, slash + 1);
        authIdent.User = (USHORT *)pszUserName;
        authIdent.UserLength = wcslen(pszUserName);

        StringCchCopyN(pszDomain, CREDUI_MAX_USERNAME_LENGTH + 1, pszName, slash - pszName);
        authIdent.Domain = (USHORT *)pszDomain;
        authIdent.DomainLength = slash - pszName;
        authIdent.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

        userAcct = &authIdent;
    }

    // Step 6: Set security levels on a WMI connection 
    hres = CoSetProxyBlanket(
        pSvc,                           // Indicates the proxy to set
        RPC_C_AUTHN_DEFAULT,            // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_DEFAULT,            // RPC_C_AUTHZ_xxx
        COLE_DEFAULT_PRINCIPAL,         // Server principal name 
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,  // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE,    // RPC_C_IMP_LEVEL_xxx
        userAcct,                       // client identity
        EOAC_NONE                       // proxy capabilities 
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 7: Use the IWbemServices pointer to make requests of WMI 

    // For example, get the name of the operating system
    IEnumWbemClassObject * pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
                           bstr_t("Select * from Win32_OperatingSystem"),
                           WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                           NULL,
                           &pEnumerator);
    if (FAILED(hres)) {
        cout << "Query for operating system name failed." << " Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 8: Secure the enumerator proxy
    hres = CoSetProxyBlanket(
        pEnumerator,                    // Indicates the proxy to set
        RPC_C_AUTHN_DEFAULT,            // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_DEFAULT,            // RPC_C_AUTHZ_xxx
        COLE_DEFAULT_PRINCIPAL,         // Server principal name 
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,  // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE,    // RPC_C_IMP_LEVEL_xxx
        userAcct,                       // client identity
        EOAC_NONE                       // proxy capabilities 
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket on enumerator. Error code = 0x" << hex << hres << endl;
        pEnumerator->Release();
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // When you have finished using the credentials,
    // erase them from memory.
    SecureZeroMemory(pszName, sizeof(pszName));
    SecureZeroMemory(pszPwd, sizeof(pszPwd));
    SecureZeroMemory(pszUserName, sizeof(pszUserName));
    SecureZeroMemory(pszDomain, sizeof(pszDomain));

    // Step 9: Get the data from the query in step 7
    IWbemClassObject * pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " OS Name : " << vtProp.bstrVal << endl;

        // Get the value of the FreePhysicalMemory property
        hr = pclsObj->Get(L"FreePhysicalMemory", 0, &vtProp, 0, 0);
        wcout << " Free physical memory (in kilobytes): " << vtProp.uintVal << endl;
        VariantClear(&vtProp);

        pclsObj->Release();
        pclsObj = NULL;
    }

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    if (pclsObj) {
        pclsObj->Release();
    }

    CoUninitialize();

    return 0;   // Program successfully completed.
}


int GettingDatafromLocalComputer(int argc, char ** argv)
/*
Example: Getting WMI Data from the Local Computer
2018/05/31

You can use the procedure and code examples in this topic to create a complete WMI client application that performs COM initialization,
connects to WMI on the local computer, retrieves data semisynchronously, and then cleans up.
This example gets the name of the operating system on the local computer and displays it.
To retrieve data from a remote computer, see Example: Getting WMI Data from a Remote Computer.
For getting the data asynchronously, see Example: Getting WMI Data from the Local Computer Asynchronously.

The following procedure is used to execute the WMI application.
Steps 1 through 5 contain all the steps required to set up and
connect to WMI, and steps 6 and 7 are where data is queried and received.

Initialize COM parameters with a call to CoInitializeEx.

For more information, see Initializing COM for a WMI Application.

Initialize COM process security by calling CoInitializeSecurity.

For more information, see Setting the Default Process Security Level Using C++.

Obtain the initial locator to WMI by calling CoCreateInstance.

For more information, see Creating a Connection to a WMI Namespace.

Obtain a pointer to IWbemServices for the root\cimv2 namespace on the local computer by calling IWbemLocator::ConnectServer.
To connect to a remote computer, see Example: Getting WMI Data from a Remote Computer.

For more information, see Creating a Connection to a WMI Namespace.

Set IWbemServices proxy security so the WMI service can impersonate the client by calling CoSetProxyBlanket.

For more information, see Setting the Security Levels on a WMI Connection.

Use the IWbemServices pointer to make requests of WMI.
This example executes a query for the name of the operating system by calling IWbemServices::ExecQuery.

The following WQL query is one of the method arguments.

SELECT * FROM Win32_OperatingSystem

The result of this query is stored in an IEnumWbemClassObject pointer.
This allows the data objects from the query to be retrieved semisynchronously with the IEnumWbemClassObject interface.
For more information, see Enumerating WMI.
For getting the data asynchronously,
see Example: Getting WMI Data from the Local Computer Asynchronously.

For more information about making requests to WMI,
see Manipulating Class and Instance Information, Querying WMI, and Calling a Method.

Get and display the data from the WQL query.
The IEnumWbemClassObject pointer is linked to the data objects that the query returned,
and the data objects can be retrieved with the IEnumWbemClassObject::Next method.
This method links the data objects to an IWbemClassObject pointer that is passed into the method.
Use the IWbemClassObject::Get method to get the desired information from the data objects.

The following code example is used to get the Name property from the data object,
which provides the name of the operating system.

C++

VARIANT vtProp;

// Get the value of the Name property
hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
After the value of the Name property is stored in the VARIANT variable vtProp,
it can then be displayed to the user.

For more information, see Enumerating WMI.

The following code example retrieves WMI data semisynchronously from a local computer.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example--getting-wmi-data-from-the-local-computer
*/
{
    HRESULT hres;

    // Step 1: Initialize COM. 
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Step 2: Set general COM security levels 
    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                    // Program has failed.
    }

    // Step 3: Obtain the initial locator to WMI 
    IWbemLocator * pLoc = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator,
                            0,
                            CLSCTX_INPROC_SERVER,
                            IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object." << " Err code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: Connect to WMI through the IWbemLocator::ConnectServer method

    IWbemServices * pSvc = NULL;

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc to make IWbemServices calls.
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
        NULL,                    // User name. NULL = current user
        NULL,                    // User password. NULL = current
        0,                       // Locale. NULL indicates current
        NULL,                    // Security flags.
        0,                       // Authority (for example, Kerberos)
        0,                       // Context object 
        &pSvc                    // pointer to IWbemServices proxy
    );
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Step 5: Set security levels on the proxy 
    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 6: Use the IWbemServices pointer to make requests of WMI 

    // For example, get the name of the operating system
    IEnumWbemClassObject * pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
                           bstr_t("SELECT * FROM Win32_OperatingSystem"),
                           WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                           NULL,
                           &pEnumerator);
    if (FAILED(hres)) {
        cout << "Query for operating system name failed." << " Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 7: Get the data from the query in step 6 
    IWbemClassObject * pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " OS Name : " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();

    return 0;   // Program successfully completed.
}


int CreatingWMIApplication(int argc, char ** argv)
/*
Example: Creating a WMI Application
2018/05/31

You can use the procedure and code examples in this topic to create a complete WMI client application that performs COM initialization,
connects to WMI on the local computer, reads some data, and cleans up.
Connecting to WMI on a Remote Computer describes how to obtain data from remote computers.

This following procedure includes all of the steps that are required by all C++ WMI applications.

Initialize COM parameters with a call to CoInitializeEx.

For more information, see Initializing COM for a WMI Application.

Initialize COM process security by calling CoInitializeSecurity.

For more information, see Setting the Default Process Security Level Using C++.

Obtain a pointer to IWbemServices for a namespace on a specified host computer—the local computer in the simple case—by calling IWbemLocator::ConnectServer.

To connect to a remote computer, for example Computer_A, use the following object path parameter:

_bstr_t(L"\\COMPUTER_A\ROOT\\CIMV2")
For more information, see Creating a Connection to a WMI Namespace.

Set the IWbemServices proxy security so WMI service can impersonate the client by calling CoSetProxyBlanket.

For more information, see Setting the Security Levels on a WMI Connection.

Use the IWbemServices pointer to make requests of WMI.
For example, querying for all the Win32_Service instances to determine which services are stopped.

For more information, see Manipulating Class and Instance Information, Querying WMI, and Receiving a WMI Event.

Clean up objects and COM.

For more information, see Cleaning up and Shutting Down a WMI Application.

The following example code is a complete WMI client application.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/example-creating-a-wmi-application
*/
{
    HRESULT hres;

    // Initialize COM.
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. " << "Error code = 0x" << hex << hres << endl;
        return 1;              // Program has failed.
    }

    // Initialize 
    hres = CoInitializeSecurity(
        NULL,
        -1,      // COM negotiates service                  
        NULL,    // Authentication services
        NULL,    // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
        RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
        NULL,             // Authentication info 
        EOAC_NONE,        // Additional capabilities
        NULL              // Reserved
    );
    if (FAILED(hres)) {
        cout << "Failed to initialize security. " << "Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;          // Program has failed.
    }

    // Obtain the initial locator to Windows Management on a particular host computer.
    IWbemLocator * pLoc = 0;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object. " << "Error code = 0x" << hex << hres << endl;
        CoUninitialize();
        return 1;       // Program has failed.
    }

    IWbemServices * pSvc = 0;

    // Connect to the root\cimv2 namespace with the
    // current user and obtain pointer pSvc to make IWbemServices calls.

    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"), // WMI namespace
        NULL,                    // User name
        NULL,                    // User password
        0,                       // Locale
        NULL,                    // Security flags                 
        0,                       // Authority       
        0,                       // Context object
        &pSvc                    // IWbemServices proxy
    );
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x" << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Set the IWbemServices proxy so that impersonation of the user (client) occurs.
    hres = CoSetProxyBlanket(
        pSvc,                         // the proxy to set
        RPC_C_AUTHN_WINNT,            // authentication service
        RPC_C_AUTHZ_NONE,             // authorization service
        NULL,                         // Server principal name
        RPC_C_AUTHN_LEVEL_CALL,       // authentication level
        RPC_C_IMP_LEVEL_IMPERSONATE,  // impersonation level
        NULL,                         // client identity 
        EOAC_NONE                     // proxy capabilities     
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Use the IWbemServices pointer to make requests of WMI. 
    // Make requests here:

    // For example, query for all the running processes
    IEnumWbemClassObject * pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
                           bstr_t("SELECT * FROM Win32_Process"),
                           WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                           NULL,
                           &pEnumerator);
    if (FAILED(hres)) {
        cout << "Query for processes failed. " << "Error code = 0x" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    } else {
        IWbemClassObject * pclsObj;
        ULONG uReturn = 0;

        while (pEnumerator) {
            hres = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
            if (0 == uReturn) {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            wcout << "Process Name : " << vtProp.bstrVal << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            pclsObj = NULL;
        }
    }

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();

    CoUninitialize();

    return 0;   // Program successfully completed.
}


int __cdecl AccessingPerformanceData(int argc, wchar_t * argv[])
/*
Accessing Performance Data in C++

The following C++ code example enumerates a high-performance class,
where the client retrieves a property handle from the first object,
and reuses the handle for the remainder of the refresh operation.
Each call to the Refresh method updates the number of instances and the instance data.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/accessing-performance-data-in-c--
*/
{
    // To add error checking, check returned HRESULT below where collected.
    HRESULT                 hr = S_OK;
    IWbemRefresher * pRefresher = NULL;
    IWbemConfigureRefresher * pConfig = NULL;
    IWbemHiPerfEnum * pEnum = NULL;
    IWbemServices * pNameSpace = NULL;
    IWbemLocator * pWbemLocator = NULL;
    IWbemObjectAccess ** apEnumAccess = NULL;
    BSTR                    bstrNameSpace = NULL;
    long                    lID = 0;
    long                    lVirtualBytesHandle = 0;
    long                    lIDProcessHandle = 0;
    DWORD                   dwVirtualBytes = 0;
    DWORD                   dwProcessId = 0;
    DWORD                   dwNumObjects = 0;
    DWORD                   dwNumReturned = 0;
    DWORD                   dwIDProcess = 0;
    DWORD                   i = 0;
    int                     x = 0;

    if (FAILED(hr = CoInitializeEx(NULL, COINIT_MULTITHREADED))) {
        goto CLEANUP;
    }

    if (FAILED(hr = CoInitializeSecurity(NULL,
                                         -1,
                                         NULL,
                                         NULL,
                                         RPC_C_AUTHN_LEVEL_NONE,
                                         RPC_C_IMP_LEVEL_IMPERSONATE,
                                         NULL, EOAC_NONE, 0))) {
        goto CLEANUP;
    }

    if (FAILED(hr = CoCreateInstance(CLSID_WbemLocator,
                                     NULL,
                                     CLSCTX_INPROC_SERVER,
                                     IID_IWbemLocator,
                                     (void **)&pWbemLocator))) {
        goto CLEANUP;
    }

    // Connect to the desired namespace.
    bstrNameSpace = SysAllocString(L"\\\\.\\root\\cimv2");
    if (NULL == bstrNameSpace) {
        hr = E_OUTOFMEMORY;
        goto CLEANUP;
    }
    if (FAILED(hr = pWbemLocator->ConnectServer(bstrNameSpace,
                                                NULL, // User name
                                                NULL, // Password
                                                NULL, // Locale
                                                0L,   // Security flags
                                                NULL, // Authority
                                                NULL, // Wbem context
                                                &pNameSpace))) {
        goto CLEANUP;
    }
    pWbemLocator->Release();
    pWbemLocator = NULL;
    SysFreeString(bstrNameSpace);
    bstrNameSpace = NULL;

    if (FAILED(hr = CoCreateInstance(CLSID_WbemRefresher,
                                     NULL,
                                     CLSCTX_INPROC_SERVER,
                                     IID_IWbemRefresher,
                                     (void **)&pRefresher))) {
        goto CLEANUP;
    }

    if (FAILED(hr = pRefresher->QueryInterface(IID_IWbemConfigureRefresher, (void **)&pConfig))) {
        goto CLEANUP;
    }

    // Add an enumerator to the refresher.
    if (FAILED(hr = pConfig->AddEnum(pNameSpace,
                                     L"Win32_PerfRawData_PerfProc_Process",
                                     0,
                                     NULL,
                                     &pEnum,
                                     &lID))) {
        goto CLEANUP;
    }
    pConfig->Release();
    pConfig = NULL;

    // Get a property handle for the VirtualBytes property.

    // Refresh the object ten times and retrieve the value.
    for (x = 0; x < 10; x++) {
        dwNumReturned = 0;
        dwIDProcess = 0;
        dwNumObjects = 0;

        if (FAILED(hr = pRefresher->Refresh(0L))) {
            goto CLEANUP;
        }

        hr = pEnum->GetObjects(0L, dwNumObjects, apEnumAccess, &dwNumReturned);
        // If the buffer was not big enough, allocate a bigger buffer and retry.
        if (hr == WBEM_E_BUFFER_TOO_SMALL && dwNumReturned > dwNumObjects) {
            apEnumAccess = new IWbemObjectAccess * [dwNumReturned];
            if (NULL == apEnumAccess) {
                hr = E_OUTOFMEMORY;
                goto CLEANUP;
            }
            SecureZeroMemory(apEnumAccess, dwNumReturned * sizeof(IWbemObjectAccess *));
            dwNumObjects = dwNumReturned;

            if (FAILED(hr = pEnum->GetObjects(0L, dwNumObjects, apEnumAccess, &dwNumReturned))) {
                goto CLEANUP;
            }
        } else {
            if (hr == WBEM_S_NO_ERROR) {
                hr = WBEM_E_NOT_FOUND;
                goto CLEANUP;
            }
        }

        // First time through, get the handles.
        if (0 == x) {
            CIMTYPE VirtualBytesType;
            CIMTYPE ProcessHandleType;
            if (FAILED(hr = apEnumAccess[0]->GetPropertyHandle(L"VirtualBytes",
                                                               &VirtualBytesType,
                                                               &lVirtualBytesHandle))) {
                goto CLEANUP;
            }
            if (FAILED(hr = apEnumAccess[0]->GetPropertyHandle(L"IDProcess",
                                                               &ProcessHandleType,
                                                               &lIDProcessHandle))) {
                goto CLEANUP;
            }
        }

        for (i = 0; i < dwNumReturned; i++) {
            if (FAILED(hr = apEnumAccess[i]->ReadDWORD(lVirtualBytesHandle, &dwVirtualBytes))) {
                goto CLEANUP;
            }
            if (FAILED(hr = apEnumAccess[i]->ReadDWORD(lIDProcessHandle, &dwIDProcess))) {
                goto CLEANUP;
            }

            wprintf(L"Process ID %lu is using %lu bytes\n", dwIDProcess, dwVirtualBytes);

            // Done with the object
            apEnumAccess[i]->Release();
            apEnumAccess[i] = NULL;
        }

        if (NULL != apEnumAccess) {
            delete[] apEnumAccess;
            apEnumAccess = NULL;
        }

        Sleep(1000);// Sleep for a second.
    }

    // exit loop here
CLEANUP:

    if (NULL != bstrNameSpace) {
        SysFreeString(bstrNameSpace);
    }

    if (NULL != apEnumAccess) {
        for (i = 0; i < dwNumReturned; i++) {
            if (apEnumAccess[i] != NULL) {
                apEnumAccess[i]->Release();
                apEnumAccess[i] = NULL;
            }
        }
        delete[] apEnumAccess;
    }
    if (NULL != pWbemLocator) {
        pWbemLocator->Release();
    }
    if (NULL != pNameSpace) {
        pNameSpace->Release();
    }
    if (NULL != pEnum) {
        pEnum->Release();
    }
    if (NULL != pConfig) {
        pConfig->Release();
    }
    if (NULL != pRefresher) {
        pRefresher->Release();
    }

    CoUninitialize();

    if (FAILED(hr)) {
        wprintf(L"Error status=%08x\n", hr);
    }

    return 1;
}


HRESULT HighPerformance()
/*
Accessing WMI Preinstalled Performance Classes

The following C++ example shows how to use the WMI high-performance API.

https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/accessing-wmi-preinstalled-performance-classes
*/
{
    // Get the local locator object
    IWbemServices * pNameSpace = NULL;
    IWbemLocator * pWbemLocator = NULL;
    CIMTYPE variant;
    //VARIANT VT;

    CoCreateInstance(CLSID_WbemLocator,
                     NULL,
                     CLSCTX_INPROC_SERVER,
                     IID_IWbemLocator,
                     (void **)&pWbemLocator);

    // Connect to the desired namespace
    BSTR bstrNameSpace = SysAllocString(L"\\\\.\\root\\cimv2");
    HRESULT hr = WBEM_S_NO_ERROR;
    hr = pWbemLocator->ConnectServer(
        bstrNameSpace,      // Namespace name
        NULL,               // User name
        NULL,               // Password
        NULL,               // Locale
        0L,                 // Security flags
        NULL,               // Authority
        NULL,               // Wbem context
        &pNameSpace         // Namespace
    );
    if (SUCCEEDED(hr)) {
        // Set namespace security.
        IUnknown * pUnk = NULL;
        pNameSpace->QueryInterface(IID_IUnknown, (void **)&pUnk);

        hr = CoSetProxyBlanket(pNameSpace,
                               RPC_C_AUTHN_WINNT,
                               RPC_C_AUTHZ_NONE,
                               NULL,
                               RPC_C_AUTHN_LEVEL_DEFAULT,
                               RPC_C_IMP_LEVEL_IMPERSONATE,
                               NULL,
                               EOAC_NONE);
        if (FAILED(hr)) {
            cout << "Cannot set proxy blanket. Error code: 0x" << hex << hr << endl;
            pNameSpace->Release();
            return hr;
        }

        hr = CoSetProxyBlanket(pUnk,
                               RPC_C_AUTHN_WINNT,
                               RPC_C_AUTHZ_NONE,
                               NULL,
                               RPC_C_AUTHN_LEVEL_DEFAULT,
                               RPC_C_IMP_LEVEL_IMPERSONATE,
                               NULL,
                               EOAC_NONE);
        if (FAILED(hr)) {
            cout << "Cannot set proxy blanket. Error code: 0x" << hex << hr << endl;
            pUnk->Release();
            return hr;
        }

        pUnk->Release();// Clean up the IUnknown.

        IWbemRefresher * pRefresher = NULL;
        IWbemConfigureRefresher * pConfig = NULL;

        // Create a WMI Refresher and get a pointer to the IWbemConfigureRefresher interface.
        CoCreateInstance(CLSID_WbemRefresher,
                         NULL,
                         CLSCTX_INPROC_SERVER,
                         IID_IWbemRefresher,
                         (void **)&pRefresher);

        pRefresher->QueryInterface(IID_IWbemConfigureRefresher, (void **)&pConfig);

        IWbemClassObject * pObj = NULL;

        // Add the instance to be refreshed.
        pConfig->AddObjectByPath(pNameSpace,
                                 L"Win32_PerfRawData_PerfProc_Process.Name=\"WINWORD\"",
                                 0L,
                                 NULL,
                                 &pObj,
                                 NULL);
        if (FAILED(hr)) {
            cout << "Cannot add object. Error code: 0x" << hex << hr << endl;
            pNameSpace->Release();
            return hr;
        }

        // For quick property retrieval, use IWbemObjectAccess.
        IWbemObjectAccess * pAcc = NULL;
        pObj->QueryInterface(IID_IWbemObjectAccess, (void **)&pAcc);

        pObj->Release();// This is not required.

        // Get a property handle for the VirtualBytes property.
        long lVirtualBytesHandle = 0;
        DWORD dwVirtualBytes = 0;

        pAcc->GetPropertyHandle(L"VirtualBytes", &variant, &lVirtualBytesHandle);

        // Refresh the object ten times and retrieve the value.
        for (int x = 0; x < 10; x++) {
            pRefresher->Refresh(0L);
            pAcc->ReadDWORD(lVirtualBytesHandle, &dwVirtualBytes);
            printf("Process is using %lu bytes\n", dwVirtualBytes);
            Sleep(1000);// Sleep for a second.
        }

        // Clean up all the objects.
        pAcc->Release();

        // Done with these too.
        pConfig->Release();
        pRefresher->Release();
        pNameSpace->Release();
    }

    SysFreeString(bstrNameSpace);
    pWbemLocator->Release();
    return hr;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
