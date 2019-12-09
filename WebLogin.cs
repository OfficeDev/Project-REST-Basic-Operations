/*
This file is based on or incorporates material from the projects listed below (Third Party OSS). The original copyright notice and the license under which Microsoft received such Third Party OSS, are set forth below. Such licenses and notices are provided for informational purposes only. Microsoft licenses the Third Party OSS to you under the licensing terms for the Microsoft product or service. Microsoft reserves all other rights not expressly granted under this agreement, whether by implication, estoppel or otherwise.  

https://github.com/SharePoint/PnP-Sites-Core/blob/master/Core/OfficeDevPnP.Core/AuthenticationManager.cs 
https://github.com/SharePoint/PnP-Sites-Core/blob/master/Core/OfficeDevPnP.Core/Utilities/CookieReader.cs

PnP-Sites-Core
Copyright (c) Microsoft Corporation
All rights reserved.
 
Provided for Informational Purposes Only   MIT License  
 
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the Software), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:  
 
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. 
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
 */
using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WebLogin
{
    /// <summary>
    /// Provides functionality for getting cookie using WebBrowser control.
    /// </summary>
    public static class WebLogin
    {
        /// <summary>
        /// Returns a SharePoint on-premises / SharePoint Online cookie. Requires claims based authentication with FedAuth cookie.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <returns>cookie string</returns>
        public static string GetWebLoginCookie(Uri siteUri)
        {
            WinINetWrapper.SupressCookiePersist();

            var authCookiesContainer = new CookieContainer();
            string authCookie = null;

            var thread = new Thread(() =>
            {
                Form form = new Form();
                WebBrowser browser = new WebBrowser
                {
                    ScriptErrorsSuppressed = false,
                    Dock = DockStyle.Fill
                };

                form.SuspendLayout();
                form.Width = 900;
                form.Height = 500;
                form.Text = "Log in to " + siteUri.ToString();
                form.Controls.Add(browser);
                form.ResumeLayout(false);

                browser.Navigated += (sender, args) =>
                {
                    if (siteUri.Host.Equals(args.Url.Host))
                    {
                        // look for FedAuth cookie
                        authCookie = WinINetWrapper.GetCookie(siteUri.ToString());
                        if (!String.IsNullOrEmpty(authCookie) && authCookie.Contains("FedAuth"))
                        {
                            authCookie = authCookie.Replace("; ", ",").Replace(";", ",");

                            // browse to about:blank before closing the window as a workaround to the leak
                            browser.Navigate("about:blank");
                        }
                    }
                    else if (args.Url.OriginalString.Equals("about:blank", StringComparison.InvariantCultureIgnoreCase))
                    {
                        form.Close();
                    }
                };

                browser.Navigate(siteUri);

                form.Focus();
                form.ShowDialog();

                browser.Dispose();
                form.Dispose();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            WinINetWrapper.SupressCookiePersistReset();
            WinINetWrapper.ClearSession();

            return authCookie;
        }
    }

    /// <summary>
    /// WinInet.dll wrapper
    /// </summary>
    internal static class WinINetWrapper
    {
        /// <summary>
        /// Enables the retrieval of cookies that are marked as "HTTPOnly". 
        /// Do not use this flag if you expose a scriptable interface, 
        /// because this has security implications. It is imperative that 
        /// you use this flag only if you can guarantee that you will never 
        /// expose the cookie to third-party code by way of an 
        /// extensibility mechanism you provide. 
        /// Version:  Requires Internet Explorer 8.0 or later.
        /// </summary>
        private const int INTERNET_COOKIE_HTTPONLY = 0x00002000;

        /// <summary>
        /// A general purpose option that is used to suppress behaviors on a process-wide basis.
        /// The lpBuffer parameter of the function must be a pointer to a DWORD containing the specific behavior to suppress.
        /// </summary>
        private const int INTERNET_OPTION_SUPPRESS_BEHAVIOR = 81;

        /// <summary>
        /// Suppresses the persistence of cookies, even if the server has specified them as persistent.
        /// Version:  Requires Internet Explorer 8.0 or later.
        /// </summary>
        private const int INTERNET_SUPPRESS_COOKIE_PERSIST = 3;

        /// <summary>
        /// Disables the INTERNET_SUPPRESS_COOKIE_PERSIST suppression, re-enabling the persistence of cookies.
        /// Any previously suppressed cookies will not become persistent. 
        /// Version:  Requires Internet Explorer 8.0 or later.
        /// </summary>
        private const int INTERNET_SUPPRESS_COOKIE_PERSIST_RESET = 4;

        /// <summary>
        /// Flushes entries not in use from the password cache on the hard disk drive.
        /// Also resets the cache time used when the synchronization mode is once-per-session.
        /// No buffer is required for this option.
        /// This is used by InternetSetOption.
        /// </summary>
        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;

        /// <summary>
        /// Returns cookie contents as a string
        /// </summary>
        /// <param name="url">Url to get cookie</param>
        /// <returns>Returns Cookie contents as a string</returns>
        public static string GetCookie(string url)
        {

            int size = 1024;
            StringBuilder sb = new StringBuilder(size);
            if (!NativeMethods.InternetGetCookieEx(url, null, sb, ref size, INTERNET_COOKIE_HTTPONLY, IntPtr.Zero))
            {
                if (size < 0)
                {
                    return null;
                }
                sb = new StringBuilder(size);
                if (!NativeMethods.InternetGetCookieEx(url, null, sb, ref size, INTERNET_COOKIE_HTTPONLY, IntPtr.Zero))
                {
                    return null;
                }
            }
            return sb.ToString();
        }

        public static void ClearSession()
        {
            SetOption(INTERNET_OPTION_END_BROWSER_SESSION, null);
        }

        public static bool SupressCookiePersist()
        {
            return SetOption(INTERNET_OPTION_SUPPRESS_BEHAVIOR, INTERNET_SUPPRESS_COOKIE_PERSIST);
        }

        public static bool SupressCookiePersistReset()
        {
            return SetOption(INTERNET_OPTION_SUPPRESS_BEHAVIOR, INTERNET_SUPPRESS_COOKIE_PERSIST_RESET);
        }

        private static bool SetOption(int settingCode, int? option)
        {
            IntPtr optionPtr = IntPtr.Zero;
            int size = 0;
            if (option.HasValue)
            {
                size = sizeof(int);
                optionPtr = Marshal.AllocCoTaskMem(size);
                Marshal.WriteInt32(optionPtr, option.Value);
            }

            bool success = NativeMethods.InternetSetOption(0, settingCode, optionPtr, size);

            if (optionPtr != IntPtr.Zero)
            {
                Marshal.Release(optionPtr);
            }

            return success;
        }

        private static class NativeMethods
        {

            [DllImport("wininet.dll", EntryPoint = "InternetGetCookieEx", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern bool InternetGetCookieEx(
                string url,
                string cookieName,
                StringBuilder cookieData,
                ref int size,
                int flags,
                IntPtr pReserved);

            [DllImport("wininet.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern bool InternetSetOption(
                int hInternet,
                int dwOption,
                IntPtr lpBuffer,
                int dwBufferLength
            );
        }
    }
}
