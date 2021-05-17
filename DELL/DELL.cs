using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services;

namespace DELL
{
    public class MModel : WebService
    {
        protected const int SleepTime = 10;
        protected const int WaitTimeOut = 10;

        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        protected static extern IntPtr FindWindowEx(
          IntPtr parentHandle,
          IntPtr childAfter,
          string lclassName,
          string windowTitle);

        [DllImport("user32.dll")]
        protected static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll")]
        protected static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        protected static extern IntPtr SendMessage(
          IntPtr hWnd,
          uint Msg,
          IntPtr wParam,
          [Out] StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        protected static extern IntPtr SendMessage(
          IntPtr hWnd,
          uint Msg,
          int wParam,
          StringBuilder lParam);

        protected void Doclick(IntPtr PROPOBJ)
        {
            MModel.SendMessage(PROPOBJ, 513U, 0U, 0);
            MModel.SendMessage(PROPOBJ, 514U, 0U, 0);
        }

        protected string GetWindowTextRaw(IntPtr hwnd)
        {
            StringBuilder lParam = new StringBuilder(MModel.SendMessage(hwnd, 14U, 0U, 0) + 1);
            MModel.SendMessage(hwnd, 13U, (IntPtr)lParam.Capacity, lParam);
            return lParam.ToString();
        }

        protected void SetWindowTextRaw(IntPtr hwnd, string text) => MModel.SendMessage(hwnd, 12U, 0, new StringBuilder(text));
    }
}
