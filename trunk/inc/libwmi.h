#pragma once

//#define _WIN32_WINNT 0x0501
#define WIN32_LEAN_AND_MEAN             // �� Windows ͷ�ļ����ų�����ʹ�õ�����
#define _WINSOCK_DEPRECATED_NO_WARNINGS
#define _WIN32_DCOM
//#define SECURITY_WIN32
#define _WIN32_DCOM

//#pragma warning(disable:4005) //���ض���

// Windows ͷ�ļ�
#include <Winsock2.h>
#include <windows.h>
#include <strsafe.h>
#include <assert.h>
#include <crtdbg.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <winioctl.h>
#include <string.h>
#include <fltuser.h>
#include <locale.h>
#include <Lmserver.h>
#include <stdarg.h>
#include <wincrypt.h>
#include <intrin.h>
#include <TlHelp32.h>
#include <aclapi.h>
#include <ShlDisp.h>
#include <Shlobj.h>
#include <Softpub.h>
#include <mscat.h>
#include <WinUser.h>
#include <direct.h>
#include <sddl.h>
#include <ws2tcpip.h>
#include <fwpsu.h>
#include <atlbase.h>
#include <mbnapi.h>
#include <iostream>
#include <netfw.h>
#include <objbase.h>
#include <oleauto.h>
#include <atlconv.h>
#define _WS2DEF_
#include <mstcpip.h>
#include <Intshcut.h>
#include <atlstr.h>
#include <comutil.h>
#include <dbt.h>
#include <comdef.h>
#include <conio.h>
#include <AccCtrl.h>
#include <Mq.h>
#include <wincred.h>
#include <lmcons.h>
#include <netioapi.h>

//�����ں���ص�ͷ�ļ���
/*
��Ϊ�⼸���ļ��г�ͻ������Ҫ��������ͻ��Ҫ������
1.����ͻ��
2.����ͬһ���ļ��а������ǡ�
3.Ҳ���ٵݹ�ģ������İ����а������ǡ�
4.ֻ����������ͻ���������ǵĹ��ܲ��ܺϲ�һ���ļ���Ҫ�ֿ��ڶ���ļ��С�
5.���ԣ����ﲻ�ܰ��������е�����һ����
6.������ͷ�ļ��а������ǡ�
7.���ľ�ɱ����ʵ���ļ��а������ǡ�

ע������5����ͬʱ��������������������κ�һ��ͬʱ������
*/
//#include <winioctl.h>
//#include <ntstatus.h> //��winnt.h��������ظ���ע�⿪�أ�WIN32_NO_STATUS
//#include <winnt.h>
//#include <NTSecAPI.h>
//#include <LsaLookup.h>
//#include <SubAuth.h> //SubAuth.h��NTSecAPI.h���ظ�����Ľṹ��ע�⿪�أ�_NTDEF_
//#include <winternl.h>
//#include <ntdef.h>
//#include <SubAuth.h>

//����USB��ص�ͷ�ļ���
#include <initguid.h> //ע��ǰ��˳��
#include <usbioctl.h>
#include <usbiodef.h>
//#include <usbctypes.h>
#include <intsafe.h>
#include <specstrings.h>
#include <usb.h>
#include <usbuser.h>

#pragma comment( lib, "ole32.lib" )
#pragma comment( lib, "oleaut32.lib" )
#pragma comment(lib, "crypt32.lib")
//#pragma comment (lib,"Url.lib")
#pragma comment(lib, "wbemuuid.lib")
//#pragma comment(lib, "cmcfg32.lib")
#pragma comment(lib,"Mpr.lib")
#pragma comment(lib, "credui.lib")
#pragma comment(lib, "comsuppw.lib")

#include <Wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

#include <VersionHelpers.h>
#pragma comment(lib, "Version.lib") 

#pragma comment (lib,"Secur32.lib")

#include <shellapi.h>
#pragma comment (lib,"Shell32.lib")

#include <Rpc.h>
#pragma comment (lib,"Rpcrt4.lib")

#include <bcrypt.h>
#pragma comment (lib, "Bcrypt.lib")

#include <ncrypt.h>
#pragma comment (lib, "Ncrypt.lib")

#include <wintrust.h>
#pragma comment (lib, "wintrust.lib")

#include <Setupapi.h>
#pragma comment (lib,"Setupapi.lib")

#include <Shlwapi.h>
#pragma comment (lib,"Shlwapi.lib")

#include <DbgHelp.h>
#pragma comment (lib,"DbgHelp.lib")

#include <psapi.h>
#pragma comment(lib, "Psapi.lib")

#include <Sfc.h>
#pragma comment(lib, "Sfc.lib")

//#include <winsock.h>
#pragma comment(lib, "Ws2_32.lib")

#include <iphlpapi.h>
#pragma comment(lib, "IPHLPAPI.lib")

#include <Wtsapi32.h>
#pragma comment(lib, "Wtsapi32.lib")

#include <Userenv.h>
#pragma comment(lib, "Userenv.lib")

#include <Sensapi.h>
#pragma comment (lib,"Sensapi.lib")

#include <Wininet.h>
#pragma comment (lib,"Wininet.lib")

#include <string>
#include <list>
#include <regex>
#include <map>
#include <set>
#include <iostream>

using namespace std;


//////////////////////////////////////////////////////////////////////////////////////////////////


EXTERN_C_START


__declspec(dllimport) int WINAPI CallingProviderMethod(int iArgCnt, char ** argv);





EXTERN_C_END


//////////////////////////////////////////////////////////////////////////////////////////////////
