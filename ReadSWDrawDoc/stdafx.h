// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>


#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // ĳЩ CString ���캯��������ʽ��

#include <atlbase.h>
#include <atlstr.h>
#include <atltypes.h>
#include <afxdisp.h>

// TODO: �ڴ˴����ó�����Ҫ������ͷ�ļ�

//Import the SOLIDWORKS type library

#import "sldworks.tlb" raw_interfaces_only, raw_native_types, no_namespace, named_guids rename("PropertySheet", "ShowPropertySheet"),rename("GetOpenFileName", "SWGetOpenFileName")

//Import the SOLIDWORKS constant type library    

#import "swconst.tlb"  raw_interfaces_only, raw_native_types, no_namespace, named_guids  