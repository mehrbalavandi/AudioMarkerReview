using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Microsoft.Win32;

namespace AudioMarkerReview.Services
{
    /// <summary>
    /// 🌟 پنجره‌ی «Open Audio» که در پوشه‌ی فایلِ صوتیِ جاری باز می‌شود و همان فایل را در فهرست
    /// <b>انتخاب و تا آن اسکرول می‌کند</b> (با فوکوس روی آن). پس در یک پوشه‌ی طولانی، برای فایلِ
    /// بعدی فقط ↓ و Enter لازم است.
    ///
    /// پنجره‌ی معمولیِ WPF (Microsoft.Win32.OpenFileDialog) فقط پوشه و نامِ فایل را می‌گیرد و
    /// راهی برای اسکرول ندارد؛ برای همین این‌جا خودِ IFileOpenDialogِ ویندوز به کار رفته. بعد از
    /// باز شدنِ پوشه (OnFolderChange)، نمای فایل‌ها (IShellView) از طریقِ IServiceProvider گرفته و
    /// SelectItem با SVSI_ENSUREVISIBLE صدا زده می‌شود. چون فهرستِ پوشه ناهم‌زمان پُر می‌شود،
    /// چند بار با فاصله‌ی کوتاه امتحان می‌شود.
    ///
    /// اگر هر چیزی در این مسیر شکست بخورد (نسخه‌ی ویندوز، COM)، پنجره‌ی معمولی باز می‌شود — فقط
    /// بدونِ اسکرول.
    /// </summary>
    public static class AudioFilePicker
    {
        public const string AudioFilter = "Audio files|*.mp3;*.wav;*.m4a;*.flac|All files|*.*";

        /// <summary>مسیرِ فایلِ انتخاب‌شده، یا null اگر کاربر منصرف شد.</summary>
        public static string? Pick(Window owner, string? currentFile)
        {
            string? current = !string.IsNullOrEmpty(currentFile) && File.Exists(currentFile) ? currentFile : null;
            if (current != null)
            {
                bool shown = false;
                try
                {
                    return PickWithShellDialog(owner, current, ref shown);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"AudioFilePicker: shell dialog failed: {ex.Message}");
                    // اگر پنجره دیده و بسته شده بود، دوباره پنجره‌ی دیگری باز نکن
                    if (shown) return null;
                }
            }
            return PickWithWpfDialog(owner, current);
        }

        private static string? PickWithWpfDialog(Window owner, string? current)
        {
            var dlg = new OpenFileDialog { Filter = AudioFilter };
            if (current != null)
            {
                dlg.InitialDirectory = Path.GetDirectoryName(current);
                dlg.FileName = Path.GetFileName(current);
            }
            return dlg.ShowDialog(owner) == true ? dlg.FileName : null;
        }

        // =====================================================================
        // IFileOpenDialog + انتخابِ فایلِ جاری در نما
        // =====================================================================

        private static string? PickWithShellDialog(Window owner, string current, ref bool shown)
        {
            string folder = Path.GetDirectoryName(Path.GetFullPath(current))!;
            var dialog = (IFileDialog)new FileOpenDialogRcw();
            uint cookie = 0;
            var events = new SelectCurrentFileEvents(current);
            try
            {
                dialog.SetFileTypes(2, new[]
                {
                    new COMDLG_FILTERSPEC { pszName = "Audio files", pszSpec = "*.mp3;*.wav;*.m4a;*.flac" },
                    new COMDLG_FILTERSPEC { pszName = "All files", pszSpec = "*.*" },
                });
                dialog.SetFileTypeIndex(1);
                dialog.GetOptions(out uint options);
                dialog.SetOptions(options | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_FILEMUSTEXIST);

                SHCreateItemFromParsingName(folder, IntPtr.Zero, typeof(IShellItem).GUID, out IShellItem folderItem);
                dialog.SetFolder(folderItem);
                dialog.SetFileName(Path.GetFileName(current));
                dialog.Advise(events, out cookie);

                IntPtr hwnd = new WindowInteropHelper(owner).Handle;
                shown = true;
                int hr = dialog.Show(hwnd);
                if (hr == ERROR_CANCELLED_HR) return null;
                Marshal.ThrowExceptionForHR(hr);

                dialog.GetResult(out IShellItem result);
                result.GetDisplayName(SIGDN_FILESYSPATH, out IntPtr namePtr);
                try { return Marshal.PtrToStringUni(namePtr); }
                finally { Marshal.FreeCoTaskMem(namePtr); }
            }
            finally
            {
                events.Stop();
                if (cookie != 0) { try { dialog.Unadvise(cookie); } catch (COMException) { } }
                Marshal.FinalReleaseComObject(dialog);
            }
        }

        /// <summary>
        /// وقتی پوشه‌ی اول باز شد، فایلِ جاری را در نما انتخاب می‌کند. چون پوشه ناهم‌زمان پُر
        /// می‌شود، هر ۸۰ میلی‌ثانیه (حداکثر ۳ ثانیه) امتحان می‌کند، و یک بارِ دیگر هم کمی بعد از
        /// موفقیت (اگر فهرست بعد از انتخاب مرتب یا بازچینی شود). اگر کاربر به پوشه‌ی دیگری برود
        /// یا خودش چیزی انتخاب کند، کنار می‌کشد.
        /// </summary>
        private sealed class SelectCurrentFileEvents : IFileDialogEvents
        {
            private readonly string _file;
            private DispatcherTimer? _timer;
            private IFileDialog? _dialog;
            private int _attempts;
            private int _successes;
            private bool _finished;
            private bool _selectingOurselves;

            public SelectCurrentFileEvents(string file) => _file = Path.GetFullPath(file);

            public void Stop()
            {
                _finished = true;
                _timer?.Stop();
                _timer = null;
                _dialog = null;
            }

            public int OnFolderChange(IFileDialog pfd)
            {
                if (_finished) return S_OK;
                if (_timer != null)
                {
                    // پوشه‌ی دوم: کاربر خودش جابه‌جا شده
                    Stop();
                    return S_OK;
                }
                _dialog = pfd;
                _timer = new DispatcherTimer(DispatcherPriority.Normal) { Interval = TimeSpan.FromMilliseconds(80) };
                _timer.Tick += (_, _) => Tick();
                _timer.Start();
                return S_OK;
            }

            public int OnSelectionChange(IFileDialog pfd)
            {
                // انتخابی که خودِ کاربر کرده (نه SelectItemِ ما) = دیگر دخالت نکن
                if (!_selectingOurselves && _successes > 0) Stop();
                return S_OK;
            }

            private void Tick()
            {
                if (_finished || _dialog == null) return;
                _attempts++;
                bool ok;
                try { ok = TrySelect(_dialog, _file); }
                catch (Exception)
                {
                    ok = false;   // هر خطایی این‌جا فقط یعنی «این بار نشد»؛ پنجره نباید بشکند
                }

                if (ok) _successes++;
                if (_successes >= 1 && _timer != null && _timer.Interval.TotalMilliseconds < 300)
                    _timer.Interval = TimeSpan.FromMilliseconds(350); // یک بارِ دیگر، بعد از پُر شدنِ کاملِ فهرست
                if (_successes >= 2 || _attempts >= 40) Stop();
            }

            private bool TrySelect(IFileDialog dialog, string file)
            {
                if (dialog is not IServiceProviderCom sp) return false;
                Guid sid = SID_SFolderView;
                Guid iid = typeof(IShellView).GUID;
                if (sp.QueryService(ref sid, ref iid, out IntPtr viewPtr) != S_OK || viewPtr == IntPtr.Zero)
                {
                    // بعضی نسخه‌ها فقط IFolderView می‌دهند؛ همان شیء IShellView را هم دارد
                    Guid iidFolderView = SID_SFolderView; // IID_IFolderView == SID_SFolderView
                    if (sp.QueryService(ref sid, ref iidFolderView, out IntPtr fvPtr) != S_OK || fvPtr == IntPtr.Zero)
                        return false;
                    try
                    {
                        Guid iidView = typeof(IShellView).GUID;
                        if (Marshal.QueryInterface(fvPtr, ref iidView, out viewPtr) != S_OK || viewPtr == IntPtr.Zero)
                            return false;
                    }
                    finally { Marshal.Release(fvPtr); }
                }

                IntPtr pidl = IntPtr.Zero;
                try
                {
                    var view = (IShellView)Marshal.GetObjectForIUnknown(viewPtr);
                    if (SHParseDisplayName(file, IntPtr.Zero, out pidl, 0, out _) != S_OK || pidl == IntPtr.Zero)
                        return false;
                    IntPtr child = ILFindLastID(pidl);
                    _selectingOurselves = true;
                    try
                    {
                        return view.SelectItem(child, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED) == S_OK;
                    }
                    finally { _selectingOurselves = false; }
                }
                finally
                {
                    if (pidl != IntPtr.Zero) Marshal.FreeCoTaskMem(pidl);
                    Marshal.Release(viewPtr);
                }
            }

            public int OnFileOk(IFileDialog pfd) { Stop(); return S_OK; }
            public int OnFolderChanging(IFileDialog pfd, IShellItem psiFolder) => S_OK;
            public int OnShareViolation(IFileDialog pfd, IShellItem psi, out int pResponse) { pResponse = 0; return S_OK; }
            public int OnTypeChange(IFileDialog pfd) => S_OK;
            public int OnOverwrite(IFileDialog pfd, IShellItem psi, out int pResponse) { pResponse = 0; return S_OK; }
        }

        // =====================================================================
        // COM / Win32 تعریف‌ها (ترتیبِ متدها = ترتیبِ vtable؛ دست نزنید)
        // =====================================================================

        private const int S_OK = 0;
        private const int ERROR_CANCELLED_HR = unchecked((int)0x800704C7);
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint FOS_PATHMUSTEXIST = 0x00000800;
        private const uint FOS_FILEMUSTEXIST = 0x00001000;
        private const uint SIGDN_FILESYSPATH = 0x80058000;
        private const uint SVSI_SELECT = 0x0001;
        private const uint SVSI_DESELECTOTHERS = 0x0004;
        private const uint SVSI_ENSUREVISIBLE = 0x0008;
        private const uint SVSI_FOCUSED = 0x0010;
        private static readonly Guid SID_SFolderView = new("cde725b0-ccc9-4519-917e-325d72fab4ce");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHParseDisplayName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

        [DllImport("shell32.dll")]
        private static extern IntPtr ILFindLastID(IntPtr pidl);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct COMDLG_FILTERSPEC
        {
            [MarshalAs(UnmanagedType.LPWStr)] public string pszName;
            [MarshalAs(UnmanagedType.LPWStr)] public string pszSpec;
        }

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7"), ClassInterface(ClassInterfaceType.None)]
        private class FileOpenDialogRcw { }

        [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialog
        {
            [PreserveSig] int Show(IntPtr parent);                                   // IModalWindow
            void SetFileTypes(uint cFileTypes, [MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IFileDialogEvents pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
        }

        [ComImport, Guid("973510db-7d7f-452b-8975-74a85828d354"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialogEvents
        {
            [PreserveSig] int OnFileOk(IFileDialog pfd);
            [PreserveSig] int OnFolderChanging(IFileDialog pfd, IShellItem psiFolder);
            [PreserveSig] int OnFolderChange(IFileDialog pfd);
            [PreserveSig] int OnSelectionChange(IFileDialog pfd);
            [PreserveSig] int OnShareViolation(IFileDialog pfd, IShellItem psi, out int pResponse);
            [PreserveSig] int OnTypeChange(IFileDialog pfd);
            [PreserveSig] int OnOverwrite(IFileDialog pfd, IShellItem psi, out int pResponse);
        }

        [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        /// <summary>COMِ IServiceProvider (نه System.IServiceProvider).</summary>
        [ComImport, Guid("6d5140c1-7436-11ce-8034-00aa006009fa"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IServiceProviderCom
        {
            [PreserveSig] int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
        }

        /// <summary>IShellView (با دو متدِ IOleWindow در ابتدا). فقط SelectItem صدا زده می‌شود؛
        /// بقیه فقط جای خودشان را در vtable نگه می‌دارند.</summary>
        [ComImport, Guid("000214E3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellView
        {
            [PreserveSig] int GetWindow(out IntPtr phwnd);                 // IOleWindow
            [PreserveSig] int ContextSensitiveHelp(int fEnterMode);        // IOleWindow
            [PreserveSig] int TranslateAccelerator(IntPtr pmsg);
            [PreserveSig] int EnableModeless(int fEnable);
            [PreserveSig] int UIActivate(uint uState);
            [PreserveSig] int Refresh();
            [PreserveSig] int CreateViewWindow(IntPtr psvPrevious, IntPtr pfs, IntPtr psb, IntPtr prcView, out IntPtr phWnd);
            [PreserveSig] int DestroyViewWindow();
            [PreserveSig] int GetCurrentInfo(IntPtr pfs);
            [PreserveSig] int AddPropertySheetPages(uint dwReserved, IntPtr pfn, IntPtr lparam);
            [PreserveSig] int SaveViewState();
            [PreserveSig] int SelectItem(IntPtr pidlItem, uint uFlags);
            [PreserveSig] int GetItemObject(uint uItem, ref Guid riid, out IntPtr ppv);
        }
    }
}
