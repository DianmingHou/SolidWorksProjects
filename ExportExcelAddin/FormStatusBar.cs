using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExportExcelAddin
{
    public partial class FormStatusBar : Form
    {
        public FormStatusBar()
        {
            InitializeComponent();
        }
        public FormStatusBar(string title, string message)
        {
            InitializeComponent();
            this.Text = title;
            this.lblMessage.Text = message;
        }
        public delegate void UserCustomHandle(object sender);
        public event UserCustomHandle UserCustomEvent;
        private delegate void DelegateFunc();

        private void FormStatusBar_Load(object sender, EventArgs e)
        {
            int iReturn = RemoveXButton(this.Handle.ToInt32());
            new Thread(new ThreadStart(()=>
            {
                UserCustomEvent(this);
                this.TryClose();
            })).Start();
        }



        [System.Runtime.InteropServices.DllImport("USER32.DLL")]
        public static extern int GetSystemMenu(int hwnd, int bRevert);
        [System.Runtime.InteropServices.DllImport("USER32.DLL")]
        public static extern int RemoveMenu(int hMenu, int nPosition, int wFlags);
        public int RemoveXButton(int iHWND)
        {
            int iSysMenu;
            const int MF_BYCOMMAND = 0x400; //0x400-关闭
            iSysMenu = GetSystemMenu(this.Handle.ToInt32(), 0);
            return RemoveMenu(iSysMenu, 6, MF_BYCOMMAND);
        }


        private void TryClose()
        {
            if (this.InvokeRequired)
                this.Invoke(new DelegateFunc(TryClose));
            else
                this.Close();

        }
    }
}
