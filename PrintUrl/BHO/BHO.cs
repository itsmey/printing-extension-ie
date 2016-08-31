using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHDocVw;
using mshtml;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices.Expando;
using System.Reflection;
using System.Xml.Serialization;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace PrintUrl.BHO
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D40C654D-7C51-4EB3-95B2-1E23905C2A2D")]
    [ProgId("MyBHO.UrlPrinter")]

    public class UrlPrinterBHO : IObjectWithSite, IOleCommandTarget
    {
        IWebBrowser2 browser;
        private object site;

        void OnDocumentComplete(object pDisp, ref object URL)
        {
            if (pDisp != this.site)
                    return;

            IHTMLDocument2 doc = (IHTMLDocument2)browser.Document;

            if (LoadData("Enabled") == "on")
            {
                browser.ExecWB(SHDocVw.OLECMDID.OLECMDID_PRINT, OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER);
                //doc.parentWindow.execScript("window.print();");
                //SendKeys.SendWait("{^P}");
                //doc.parentWindow.execScript("window.close();");
                doc.parentWindow.close();

                string next_url = PopUrl();
                if (next_url != "")
                {
                    browser.Navigate(next_url, 0x800);
                    //doc.parentWindow.execScript("window.onunload = function() {window.open('" + next_url + "')};");
                }

                else
                    SetData("Enabled", "off");

            }
        }

        void OnQuit()
        {
            SetData("Urls", "");
            SetData("Enabled", "off");
        }

        void OnWindowClosing(bool IsChildWindow, ref bool Cancel)
        {
            //MessageBox.Show("closing");
        }

        void OnNavigateComplete2(object pDisp, ref object url)
        {
            //MessageBox.Show("complete");
        }

        private string PopUrl()
        {
            string url_line = LoadData("Urls");
            if (url_line == "" || url_line == null)
                return "";

            string new_url_line = "";
            string[] urls = url_line.Remove(url_line.Length - 1).Split(' ');

            for (int i = 1; i < urls.Length; i++)
                new_url_line = new_url_line + urls[i] + " ";

            SetData("Urls", new_url_line);

            return urls[0];
        }

        #region Load and Save Data
        public static string RegData = "Software\\PrintUrlExtension";

        [DllImport("ieframe.dll")]
        public static extern int IEGetWriteableHKCU(ref IntPtr phKey);

        public static void SetData(string key, object value)
        {
            IntPtr phKey = new IntPtr();
            var answer = IEGetWriteableHKCU(ref phKey);
            RegistryKey writeable_registry = RegistryKey.FromHandle(
                new Microsoft.Win32.SafeHandles.SafeRegistryHandle(phKey, true)
            );
            RegistryKey registryKey = writeable_registry.OpenSubKey(RegData, true);

            if (registryKey == null)
                registryKey = writeable_registry.CreateSubKey(RegData);
            registryKey.SetValue(key, value);

            writeable_registry.Close();
        }

        private static string LoadData(string key)
        {
            IntPtr phKey = new IntPtr();
            var answer = IEGetWriteableHKCU(ref phKey);
            RegistryKey writeable_registry = RegistryKey.FromHandle(
                new Microsoft.Win32.SafeHandles.SafeRegistryHandle(phKey, true)
            );
            RegistryKey registryKey = writeable_registry.OpenSubKey(RegData, true);

            string result;

            if (registryKey == null)
                result = "off";
            else
              result = (string)registryKey.GetValue(key);

            writeable_registry.Close();
            return result;
        }
        #endregion

        [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
        [InterfaceType(1)]
        public interface IServiceProvider
        {
            int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
        }

        #region Implementation of IObjectWithSite
        int IObjectWithSite.SetSite(object site)
        {
            this.site = site;

            if (site != null)
            {
                var serviceProv = (IServiceProvider)this.site;
                var guidIWebBrowserApp = Marshal.GenerateGuidForType(typeof(IWebBrowserApp)); // new Guid("0002DF05-0000-0000-C000-000000000046");
                var guidIWebBrowser2 = Marshal.GenerateGuidForType(typeof(IWebBrowser2)); // new Guid("D30C1661-CDAF-11D0-8A3E-00C04FC9E26E");
                IntPtr intPtr;
                serviceProv.QueryService(ref guidIWebBrowserApp, ref guidIWebBrowser2, out intPtr);

                browser = (IWebBrowser2)Marshal.GetObjectForIUnknown(intPtr);

                ((DWebBrowserEvents2_Event)browser).DocumentComplete +=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                ((DWebBrowserEvents2_Event)browser).OnQuit +=
                    new DWebBrowserEvents2_OnQuitEventHandler(this.OnQuit);
                ((DWebBrowserEvents2_Event)browser).WindowClosing +=
                    new DWebBrowserEvents2_WindowClosingEventHandler(this.OnWindowClosing);
                ((DWebBrowserEvents2_Event)browser).NavigateComplete2 +=
                    new DWebBrowserEvents2_NavigateComplete2EventHandler(this.OnNavigateComplete2);
            }
            else
            {
                ((DWebBrowserEvents2_Event)browser).DocumentComplete -=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                ((DWebBrowserEvents2_Event)browser).OnQuit +=
                    new DWebBrowserEvents2_OnQuitEventHandler(this.OnQuit);
                ((DWebBrowserEvents2_Event)browser).WindowClosing +=
                    new DWebBrowserEvents2_WindowClosingEventHandler(this.OnWindowClosing);
                ((DWebBrowserEvents2_Event)browser).NavigateComplete2 +=
                    new DWebBrowserEvents2_NavigateComplete2EventHandler(this.OnNavigateComplete2);
                browser = null;
            }
            return 0;
        }
        int IObjectWithSite.GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(browser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }
        #endregion
        #region Implementation of IOleCommandTarget
        int IOleCommandTarget.QueryStatus(IntPtr pguidCmdGroup, uint cCmds, ref OLECMD prgCmds, IntPtr pCmdText)
        {
            return 0;
        }
        int IOleCommandTarget.Exec(IntPtr pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            try
            {
                var document2 = browser.Document as IHTMLDocument2;

                string script = @"var elements = document.querySelectorAll('.to_print'); 
                                                  var urls_str = """";
                                                  for(var i=0; i<elements.length; i++) {
                                                    if (elements[i].checked == true)  {
                                                      urls_str = urls_str.concat(elements[i].getAttribute('value'));
                                                      urls_str = urls_str.concat(' ');
                                                    }
                                                  };
                                                  if (urls_str == """") alert('No URLs to print!');
                                                  document.MyUrls = urls_str;";

                document2.parentWindow.execScript(script);

                PropertyInfo property = ((IExpando)document2).GetProperty("MyUrls", BindingFlags.Default);
                if (property != null)
                {
                    object value = property.GetValue(document2, null);
                    if (value != null)
                    {
                        SetData("Urls", value.ToString());
                        string next_url = PopUrl();
                        if (next_url != "")
                        {
                            SetData("Enabled", "on");
                            browser.Navigate(next_url, 0x800);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return 0;
        }
        #endregion

        #region Registering with regasm
        public static string RegBHO = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        public static string RegCmd = "Software\\Microsoft\\Internet Explorer\\Extensions";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            SetData("Urls", "");
            SetData("Enabled", "off");

            string guid = type.GUID.ToString("B");

            // BHO
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegBHO, true);
                if (registryKey == null)
                    registryKey = Registry.LocalMachine.CreateSubKey(RegBHO);
                RegistryKey key = registryKey.OpenSubKey(guid, true);
                if (key == null)
                    key = registryKey.CreateSubKey(guid);
                key.SetValue("Alright", 1);
                registryKey.Close();
                key.Close();
            }

            // Command
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegCmd, true);
                if (registryKey == null)
                    registryKey = Registry.LocalMachine.CreateSubKey(RegCmd);
                RegistryKey key = registryKey.OpenSubKey(guid, true);
                if (key == null)
                    key = registryKey.CreateSubKey(guid);
                key.SetValue("ButtonText", "Print");
                key.SetValue("CLSID", "{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}");
                key.SetValue("ClsidExtension", guid);
                key.SetValue("Icon", "");
                key.SetValue("HotIcon", "");
                key.SetValue("Default Visible", "Yes");
                //key.SetValue("MenuText", "&Highlighter options");
                key.SetValue("ToolTip", "Print");
                //key.SetValue("KeyPath", "no");
                registryKey.Close();
                key.Close();
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            string guid = type.GUID.ToString("B");
            // BHO
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegBHO, true);
                if (registryKey != null)
                    registryKey.DeleteSubKey(guid, false);
            }
            // Command
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegCmd, true);
                if (registryKey != null)
                    registryKey.DeleteSubKey(guid, false);
            }
        }
        #endregion
    }
}
