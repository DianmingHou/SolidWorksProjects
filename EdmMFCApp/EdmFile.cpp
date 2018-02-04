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


// 是否是文件夹
bool CEdmFile::IsFolder()
{
	return m_Folder;
}


// 设置该文件为文件夹
void CEdmFile::SetFolder()
{
	m_Folder = true;
}


// 获取文件类型
CString CEdmFile::GetFileType()
{
	return m_FileType;
}


void CEdmFile::SetFileType(CString iFileType)
{
	m_FileType = iFileType;
}


// 获取文件路径
CString CEdmFile::GetFilePath()
{
	return m_FilePath;
}


// 设置文件路径
void CEdmFile::SetFilePath(CString iFilePath)
{
	m_FilePath = iFilePath;
}
