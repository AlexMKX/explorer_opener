using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;

namespace exploreropener
{

    class Opener
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [STAThread]
        static void Main(string[] args)
        {
            IntPtr handle = GetForegroundWindow();
            List<string> selected = new List<string>();
            Shell32.Shell shell = new Shell32.Shell();

            foreach (var window in shell.Windows())
            {
                var testy = window.HWND;
                var testy2 = (int)handle;
                if (window.HWND == (int)handle)
                {
                    Shell32.FolderItems items = ((Shell32.IShellFolderViewDual2)window.Document).SelectedItems();
                    foreach (Shell32.FolderItem item in items)
                    {
                        if (args[0] == "browse")
                        {
                            if (item.Path[1] == ':')
                            {
                                String s = String.Format("/select,\"{0}\"", item.Path);
                                
                                shell.ShellExecute("explorer.exe", s);
                            }
                            else
                            {
                                shell.Explore(Path.GetDirectoryName(item.Path));
                            }
                        }
                        if (args[0] == "open")
                        {
                            shell.ShellExecute(args[1],item.Path);
                        }
                    }
                }
            }
            
        }
    }
}
