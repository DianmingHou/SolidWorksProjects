#pragma once
#include "afxcmn.h"


class CEdmFile
{
public:
	CEdmFile();
	CEdmFile(CString iFileType,CString iFilePath,bool isFolder=false);
	~CEdmFile();
private:
	bool m_Folder;
	CString m_FileType;
	CString m_FilePath;
public:
	// 是否是文件夹
	bool IsFolder();
	// 设置该文件为文件夹
	void SetFolder();
	// 获取文件类型
	CString GetFileType();
	void SetFileType(CString iFileType);
	// 获取文件路径
	CString GetFilePath();
	// 设置文件路径
	void SetFilePath(CString iFilePath);
};
