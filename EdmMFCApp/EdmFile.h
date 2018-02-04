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
	// �Ƿ����ļ���
	bool IsFolder();
	// ���ø��ļ�Ϊ�ļ���
	void SetFolder();
	// ��ȡ�ļ�����
	CString GetFileType();
	void SetFileType(CString iFileType);
	// ��ȡ�ļ�·��
	CString GetFilePath();
	// �����ļ�·��
	void SetFilePath(CString iFilePath);
};
