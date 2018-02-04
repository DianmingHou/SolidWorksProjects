// ReadSWDrawDoc.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <iostream>
using namespace std;

int _tmain(int argc, _TCHAR* argv[])
{
	//Initialize COM 

	CoInitialize(NULL);  
	//Use ATL smart pointers
	CComPtr<ISldWorks> swApp; 
	HRESULT rc = S_OK;
	//Create an instance of SOLIDWORKS
	HRESULT hres = swApp.CoCreateInstance(__uuidof(SldWorks), NULL, CLSCTX_LOCAL_SERVER); 
	// Your code
	IModelDoc2Ptr m_pModelDoc = nullptr;
	long error = 0,warnning = 0;
	//Default optionconfig
	//rc = swApp->ActivateDoc3(L"D:/23.SLDDRW",VARIANT_TRUE,0,&error,(IDispatch**)&m_pModelDoc);
	rc = swApp->OpenDoc6(L"E:/Models/SOLIDWORKS/G20-A01P.SLDDRW",swDocDRAWING,swOpenDocOptions_Silent,L"",&error,&warnning,&m_pModelDoc);
	if (FAILED(rc)||error!=0||m_pModelDoc==nullptr)
	{
		cout<<" cannot open failed and error code is "<<error<<endl;
		return 1;
	}
	CComPtr<IFeature> featurePtr = nullptr;
	m_pModelDoc->FirstFeature((IDispatch**)&featurePtr);
	CComPtr<IFeature> swNextFeat;
	CString strApp;
	while (featurePtr!=nullptr)
	{
		BSTR typeName;
		//featurePtr->GetTypeName2(&typeName);
		//CString typeNameStr = typeName;
		//strApp+=L"\t" +typeNameStr;
		//featurePtr->IGetNextFeature(&swNextFeat);
		//featurePtr.Release();

		//featurePtr = swNextFeat;
		break;

	}
	cout<<strApp<<endl;
	IDrawingDocPtr m_pDrawingDoc = (IDrawingDocPtr)m_pModelDoc;
	if (FAILED(rc)||!m_pDrawingDoc)
	{
		m_pModelDoc=nullptr;
		return E_FAIL;
	}
	//m_iDrawingDoc->
	IViewPtr pView = nullptr;
	rc = m_pDrawingDoc->GetFirstView((IDispatch**)&pView);
	if (FAILED(rc)||pView==nullptr)
	{
		return E_FAIL;
	}
	while (pView!=nullptr)
	{
		long viewType = 0;
		//pView->get_Type(&viewType);
		BSTR viewName = NULL;
		pView->GetName2(&viewName);
		cout<<" view name "<<viewName<<endl;
		continue;
		if (viewType!=swDrawingSheet)
		{
			IBomTablePtr pbomTable;
			//BomFeature::GetTableAnnotations.
			pView->GetBomTable((IDispatch**)&pbomTable);
			if (pbomTable!=nullptr)
			{
				VARIANT_BOOL bool1;
				pbomTable->Attach3(&bool1);
				long rowCount =0, colCount = 0;
				pbomTable->GetTotalColumnCount(&colCount);
				pbomTable->GetTotalRowCount(&rowCount);
				for (int colIndex=0;colIndex<colCount;colIndex++)
				{
					BSTR colName;
					pbomTable->GetHeaderText(colIndex,&colName);
					CString colNameS = colName;
					strApp+=L"\t"+colNameS;
				}
				strApp+="\n";
				for (int rowIndex=0;rowIndex<rowCount;rowIndex++)
				{
					for (int colIndex=0;colIndex<colCount;colIndex++)
					{
						BSTR colName;
						pbomTable->GetEntryText(rowIndex,colIndex,&colName);
						CString context = colName;
						strApp+=L"\t"+context;
					}
					strApp+="\n";
				}
				pbomTable->Release();
			}
			IViewPtr pNextPtr;
			pView->GetNextView((IDispatch**)&pNextPtr);
			pView->Release();
			pView = pNextPtr;
		}

	}
	m_pDrawingDoc->Release();


	cout<<strApp<<endl;


	cout<<"open OK "<<endl;
	m_pModelDoc->Close();
	//Shut down SOLIDWORKS 
	swApp->ExitApp();
	//Release COM reference
	swApp = NULL;
	//Uninitialize COM
	CoUninitialize();
	return 0;
}

