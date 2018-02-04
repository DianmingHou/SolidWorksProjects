#include "stdafx.h"
#include "EdmFile.h"



#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CEdmFile::CEdmFile()
{
	m_Folder = false;
	m_FileType = _T("");
	m_FilePath = _T("");
}

CEdmFile::CEdmFile(CString iFileType, CString iFilePath,bool isFolder) {
	m_Folder = isFolder;
	m_FileType = iFileType;
	m_FilePath = iFilePath;
}

CEdmFile::~CEdmFile()
{
}


// �Ƿ����ļ���
bool CEdmFile::IsFolder()
{
	return m_Folder;
}


// ���ø��ļ�Ϊ�ļ���
void CEdmFile::SetFolder()
{
	m_Folder = true;
}


// ��ȡ�ļ�����
CString CEdmFile::GetFileType()
{
	return m_FileType;
}


void CEdmFile::SetFileType(CString iFileType)
{
	m_FileType = iFileType;
}


// ��ȡ�ļ�·��
CString CEdmFile::GetFilePath()
{
	return m_FilePath;
}


// �����ļ�·��
void CEdmFile::SetFilePath(CString iFilePath)
{
	m_FilePath = iFilePath;
}
