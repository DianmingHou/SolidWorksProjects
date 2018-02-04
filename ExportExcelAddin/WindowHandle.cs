using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportExcelAddin
{
    class WindowHandle : IWin32Window
    {
        private IntPtr mHwnd;

        public WindowHandle(int hWnd)
        {
            mHwnd = new IntPtr(hWnd);
        }
        public IntPtr Handle
        {
            get { return mHwnd; }
        }

    }
}
