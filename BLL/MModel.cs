using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services;

namespace BLL
{
    public class MModel : WebService
    {
        public const int SleepTime = 10;
        public const int WaitTimeOut = 10;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindowEx(
          IntPtr parentHandle,
          IntPtr childAfter,
          string lclassName,
          string windowTitle);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(
          IntPtr hWnd,
          uint Msg,
          IntPtr wParam,
          [Out] StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(
          IntPtr hWnd,
          uint Msg,
          int wParam,
          StringBuilder lParam);

        public void Doclick(IntPtr PROPOBJ)
        {
            MModel.SendMessage(PROPOBJ, 513U, 0U, 0);
            MModel.SendMessage(PROPOBJ, 514U, 0U, 0);
        }

        public string GetWindowTextRaw(IntPtr hwnd)
        {
            StringBuilder lParam = new StringBuilder(MModel.SendMessage(hwnd, 14U, 0U, 0) + 1);
            MModel.SendMessage(hwnd, 13U, (IntPtr)lParam.Capacity, lParam);
            return lParam.ToString();
        }

        public void SetWindowTextRaw(IntPtr hwnd, string text) => MModel.SendMessage(hwnd, 12U, 0, new StringBuilder(text));
    }
}
