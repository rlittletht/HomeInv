﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using StatusBox;
using TCore.Util;

namespace TCore.WebControl
{
    public partial class WebControl : Form
    {
        private StatusRpt m_srpt;
        public WebControl()
        {
            InitializeComponent();
        }

        public WebBrowser AxWeb { get { return m_wbc; } }

        public IHTMLDocument  Document => (IHTMLDocument)m_wbc.Document?.DomDocument;
        public IHTMLDocument2 Document2 => (IHTMLDocument2)m_wbc.Document?.DomDocument;
        public IHTMLDocument3 Document3 => (IHTMLDocument3)m_wbc.Document?.DomDocument;
        public IHTMLDocument4 Document4 => (IHTMLDocument4)m_wbc.Document?.DomDocument;
        public IHTMLDocument5 Document5 => (IHTMLDocument5)m_wbc.Document?.DomDocument;

        private async Task<bool> PageLoad(int TimeOut)
        {
            TaskCompletionSource<bool> PageLoaded = null;
            PageLoaded = new TaskCompletionSource<bool>();
            int TimeElapsed = 0;
            m_wbc.DocumentCompleted += (s, e) =>
                {
                if (m_wbc.ReadyState != WebBrowserReadyState.Complete)
                        return;
                if (PageLoaded.Task.IsCompleted)
                        return;
                PageLoaded.SetResult(true);
                };
            //
            while (PageLoaded.Task.Status != TaskStatus.RanToCompletion)
                {
                await Task.Delay(10); //interval of 10 ms worked good for me
                TimeElapsed++;
                if (TimeElapsed >= TimeOut*100)
                    {
                    PageLoaded.TrySetResult(true);
                    return false;
                    //This prevents your method or thread from waiting forever
                    }
                }
            return true;
        }
        public WebControl(StatusRpt srpt)
        {
#if notused
            m_plNewWindow3 = new List<DWebBrowserEvents2_NewWindow3EventHandler>();
            m_plBeforeNav2= new List<DWebBrowserEvents2_BeforeNavigate2EventHandler>();
#endif
            m_srpt = srpt;

            InitializeComponent();
            m_wbc.ScriptErrorsSuppressed = true;
            
        }

        bool m_fNavDone;
        private string m_sUrlExpected;
        public void ResetNav()
        {
            m_fNavDone = false;
            m_sUrlExpected = null;
        }


        async void WaitForBrowserReadyNew()
        {
            await PageLoad(10000);
        }

        void WaitForBrowserReady()
        {
            long s = 0;

            while (s < 200 && m_wbc.IsBusy)
            {
                Application.DoEvents();
                Thread.Sleep(50);
                s++;
            }

            if (s > 20000)
            {
                m_wbc.Stop();
                m_wbc.Visible = true;
                throw new Exception("browser can't get unbusy");
                
            }
            return;
        }

        
        /* R E P O R T  N A V  S T A T E */
        /*----------------------------------------------------------------------------
        	%%Function: ReportNavState
        	%%Qualified: ArbWeb.WebControl.ReportNavState
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void ReportNavState(string sTag)
        {
            m_srpt.LogData(String.Format("{0}: Busy: {1}, State: {3}, m_fNavDone: {2}", sTag, m_wbc.IsBusy, m_fNavDone, m_wbc.ReadyState), 3, StatusRpt.MSGT.Body);
        }

        public delegate bool FNavToPageDel(WebBrowser wbc, string sUrl);

        /* D O  N A V  T O  P A G E */
        /*----------------------------------------------------------------------------
        	%%Function: DoNavToPage
        	%%Qualified: ArbWeb.WebControl.DoNavToPage
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public bool DoNavToPage(WebBrowser wbc, string sUrl)
        {
            ReportNavState(String.Format("Entering FNavToPage({0})", sUrl));
            wbc.Stop();

            // m_wbc.Stop();
            WaitForBrowserReady();
            m_fNavDone = false;
            m_sUrlExpected = sUrl;

            wbc.Navigate(sUrl);
            wbc.Visible = true;

            return FWaitForNavFinish();
        }

        /* F  N A V  T O  P A G E */
        /*----------------------------------------------------------------------------
        	%%Function: FNavToPage
        	%%Qualified: ArbWeb.WebControl.FNavToPage
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public bool FNavToPage(string sUrl)
        {
            if (m_wbc.InvokeRequired)
                {
                IAsyncResult rslt = m_wbc.BeginInvoke(new FNavToPageDel(DoNavToPage), new object[] {m_wbc, sUrl});
                return (bool)m_wbc.EndInvoke(rslt);
                }
            else
                return DoNavToPage(m_wbc, sUrl);
        }

        public bool FWaitForNavFinishNew()
        {
            Task<bool> t = PageLoad(10000);

            t.Wait();
            return t.Result;
        }

        /* F  W A I T  F O R  N A V  F I N I S H */
        /*----------------------------------------------------------------------------
        	%%Function: FWaitForNavFinish
        	%%Qualified: ArbWeb.WebControl.FWaitForNavFinish
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public bool FWaitForNavFinish(string sidWaitFor = null)
        {
            long s = 0;
            
            // ok, always yield and allow it to run first (so things can get pumping)
            Application.DoEvents();
            Thread.Sleep(50);
            Application.DoEvents();

            ReportNavState("Entering WaitForNavFinish: ");
            WaitForBrowserReady();
            while (s < 200 && !m_fNavDone)
                {
                Application.DoEvents();
                if (m_wbc.ReadyState == WebBrowserReadyState.Complete)
                    {
                    if (sidWaitFor != null)
                        {
                        // we are being asked to wait for an element to exist
                        if (FCheckForControl((IHTMLDocument2) m_wbc.Document.DomDocument, sidWaitFor))
                            break;
                        }
                    else
                        {
                        break;
                        }
                    }
                Thread.Sleep(50);
                s++;
                }

            ReportNavState(String.Format("After NavDone Loop: {0}", s >= 200 ? "TIMEOUT" : "Completed"));

#if notused
            if (m_fNavIntercept)
                {
                m_wbc.Stop();
                WaitForBrowserReady();
                
                m_fNavDone = false;
                m_fNavIntercept = false;
                m_fDontIntercept = true;
                object o = m_rgbPostData;
                object sHdr = "Content-Type: application/x-www-form-urlencoded" + Environment.NewLine;
                object oHdr = sHdr;
                int flags = 0x1;
                object oFlags = flags;
                string sTgt = "_SELF";
                object oTgt = sTgt;
                
                m_wbc.Navigate(m_sNavIntercept, ref Zero, ref oTgt, ref o, ref oHdr);
                bool f = FWaitForNavFinish();
                m_fDontIntercept = false;
                
                return f;
                }
#endif                 
            if (s > 20000)
            {
                m_wbc.Stop();
                m_wbc.Visible = true;
                return false;
            }
            return true;
        }

        /* R E F R E S H  P A G E */
        /*----------------------------------------------------------------------------
        	%%Function: RefreshPage
        	%%Qualified: ArbWeb.WebControl.RefreshPage
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void RefreshPage()
        {
            m_fNavDone = false;
            m_wbc.Refresh();
            FWaitForNavFinish();
        }
        /* T R I G G E R  D O C U M E N T  D O N E */
        /*----------------------------------------------------------------------------
        	%%Function: TriggerDocumentDone
        	%%Qualified: ArbWeb.WebControl.TriggerDocumentDone
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void TriggerDocumentDone(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (m_sUrlExpected == null || String.Compare(e.Url.ToString(), m_sUrlExpected, true) == 0)
                {
                ReportNavState(String.Format("TriggerDocumentDone MATCH: {0}", e.Url));
                m_fNavDone = true;
                }
            else
                {
                ReportNavState(String.Format("TriggerDocumentDone NO MATCH: {0} {1}", e.Url, m_sUrlExpected));
                }
        }
        /* F  S E T  C H E C K B O X  C O N T R O L  V A L */
        /*----------------------------------------------------------------------------
        	%%Function: FSetCheckboxControlVal
        	%%Qualified: ArbWeb.WebControl.FSetCheckboxControlVal
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        static public bool FSetCheckboxControlVal(IHTMLDocument2 oDoc2, bool fChecked, string sName)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            
            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (ihie.@checked == fChecked)
                        return false;

                    ihie.@checked = fChecked;
                    return true;
                    }
                }
            return false;
        }

        public bool FSetTextareaControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            m_srpt.LogData(String.Format("FSetTextareaControlText for id {0}", sName), 3, StatusRpt.MSGT.Body);

            bool f = WebControl.FSetTextareaControlTextForDoc(oDoc2, sName, sValue, fCheck);

            m_srpt.LogData(String.Format("Return: {0}", f), 3, StatusRpt.MSGT.Body);
            return f;
        }
        /* F  S E T  T E X T A R E A  C O N T R O L  T E X T */
        /*----------------------------------------------------------------------------
        	%%Function: FSetTextareaControlText
        	%%Qualified: ArbWeb.WebControl.FSetTextareaControlText
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static bool FSetTextareaControlTextForDoc(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("textarea");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLTextAreaElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (fCheck)
                        {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                        }
                    else
                        {
                        fNeedSave = true;
                        }
                    ihie.value = sValue;
                    }
                }
            return fNeedSave;
        }
        /* F  S E T  I N P U T  C O N T R O L  T E X T */
        /*----------------------------------------------------------------------------
        	%%Function: FSetInputControlText
        	%%Qualified: ArbWeb.WebControl.FSetInputControlText
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static bool FSetInputControlText(IHTMLDocument2 oDoc2, string sName, string sId, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLElement ihe in hec)
                {
                IHTMLInputElement ihie = (IHTMLInputElement) ihe;

                if ((sName != null && String.Compare(ihie.name, sName, true) == 0)
                    || (sId != null && ihe.id != null && String.Compare(ihe.id, sId, true) == 0))
                    {
                    if (fCheck)
                        {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                        }
                    else
                        {
                        fNeedSave = true;
                        }
                    ihie.value = sValue;
                    }
                }
            return fNeedSave;
        }

        public static bool FSetSelectControlValue(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLSelectElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    if (fCheck)
                        {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                        }
                    else
                        {
                        fNeedSave = true;
                        }
                    ihie.value = sValue;
                    }
                }
            return fNeedSave;
        }
        /* M P  G E T  S E L E C T  V A L U E S */
		/*----------------------------------------------------------------------------
			%%Function: MpGetSelectValues
			%%Qualified: ArbWeb.AwMainForm.MpGetSelectValues
			%%Contact: rlittle

            for a given <select name=$sName><option value=$sValue>$sText</option>...
         
            Find the given sName select object. Then add a mapping of
            $sText -> $sValue to a dictionary and return it.
		----------------------------------------------------------------------------*/
        public static Dictionary<string, string> MpGetSelectValues(StatusRpt srpt, IHTMLDocument2 oDoc2, string sName)
        {
            IHTMLElementCollection hec;
			
            Dictionary<string, string> mp = new Dictionary<string, string>();

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (mp.ContainsKey(ihoe.text))
                            srpt.AddMessage(String.Format("How strange!  '{0}' shows up more than once as a position", ihoe.text), StatusRpt.MSGT.Warning);
                        else
                            mp.Add(ihoe.text, ihoe.value);
                        }
                    }
                }
            return mp;
        }

        public static List<string> ListItemsGet(IHTMLDocument2 oDoc2, string sId)
        {
            IHTMLElement ihe = (IHTMLElement) oDoc2.all.item(sId, 0);
            if (ihe == null)
                return null;

            List<string> pls = new List<string>();

            foreach (IHTMLElement iheSectionChild in (IHTMLElementCollection) ihe.children)
                {
                if (string.Compare(iheSectionChild.tagName, "ul", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                    foreach (IHTMLElement iheListItem in (IHTMLElementCollection)iheSectionChild.children)
                        {
                        if (string.Compare(iheListItem.tagName, "li", StringComparison.OrdinalIgnoreCase) == 0)
                            pls.Add(iheListItem.innerText);
                        }
                    break;
                    }
                }

            if (pls.Count == 0)
                return null;

            return pls;
        }

        // if fValueIsValue == false, then sValue is the "text" of the option control
        /* F  S E L E C T  M U L T I  S E L E C T  O P T I O N */
        /*----------------------------------------------------------------------------
        	%%Function: FSelectMultiSelectOption
        	%%Qualified: ArbWeb.WebControl.FSelectMultiSelectOption
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
            public static
            bool FSelectMultiSelectOption(IHTMLDocument2 oDoc2, string sName, string sValue, bool fValueIsValue)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");

            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if ((fValueIsValue && ihoe.value == sValue) ||
                            (!fValueIsValue && ihoe.text == sValue))
                            {
                            ihoe.selected = true;
                            return true;
                            }
                        }
                    }
                }
            return false;
        }
        /* F  R E S E T  M U L T I  S E L E C T  O P T I O N S */
        /*----------------------------------------------------------------------------
        	%%Function: FResetMultiSelectOptions
        	%%Qualified: ArbWeb.WebControl.FResetMultiSelectOptions
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static bool FResetMultiSelectOptions(IHTMLDocument2 oDoc2, string sName)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");

            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        ihoe.selected = false;
                        }
                    }
                }
            return true;
        }

        /* S  G E T  F I L T E R  I  D */
        /*----------------------------------------------------------------------------
        	%%Function: SGetFilterID
        	%%Qualified: ArbWeb.WebControl.SGetFilterID
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public string SGetFilterID(IHTMLDocument2 oDoc2, string sName, string sValue)
        {
            m_srpt.LogData(String.Format("SGetSelectIDFromDoc for id {0}", sName), 3, StatusRpt.MSGT.Body);

            string s = SGetSelectIDFromDoc(oDoc2, sName, sValue);

            m_srpt.LogData(String.Format("Return: {0}", s), 3, StatusRpt.MSGT.Body);
            return s;
        }


        /* S  G E T  F I L T E R  I  D */
        /*----------------------------------------------------------------------------
        	%%Function: SGetFilterID
        	%%Qualified: ArbWeb.WebControl.SGetFilterID
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string SGetSelectIDFromDoc(IHTMLDocument2 oDoc2, string sName, string sOptionName)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sOptionName)
                            {
                            return ihoe.value;
                            }
                        }
                    }
                }
            return null;
        }

        public static string SGetSelectValFromDoc(IHTMLDocument2 oDoc2, string sName, string sValue)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.value == sValue)
                            {
                            return ihoe.text;
                            }
                        }
                    }
                }
            return null;
        }
        public bool FSetSelectControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            m_srpt.LogData(String.Format("FSetSelectControlText for id {0}", sName), 3, StatusRpt.MSGT.Body);

            bool f = FSetSelectControlTextFromDoc(this, oDoc2, sName, sValue, fCheck);

            m_srpt.LogData(String.Format("Return: {0}", f), 3, StatusRpt.MSGT.Body);
            return f;
        }

        public void WaitDoLog(int msec)
        {
            while (msec > 0)
                {
                ReportNavState(String.Format("WaitDoLog {0}", msec));
                Thread.Sleep(50);
                msec -= 50;
                Application.DoEvents();
                }
        }

        /* F  S E T  S E L E C T  C O N T R O L  T E X T */
        /*----------------------------------------------------------------------------
        	%%Function: FSetSelectControlText
        	%%Qualified: ArbWeb.WebControl.FSetSelectControlText
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        static public bool FSetSelectControlTextFromDoc(WebControl awc, IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            bool fNeedSave = false;
            foreach (IHTMLSelectElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach(IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sValue)
                            {
                            // value is already set...
                            if (ihie.value == ihoe.value)
                                return false;
                            ihoe.selected = true;
                            object dummy = null;
                            IHTMLDocument4 oDoc4 = (IHTMLDocument4)oDoc2;
                            object eventObj = oDoc4.CreateEventObject(ref dummy);
                            HTMLSelectElementClass hsec = ihie as HTMLSelectElementClass;
                            awc.ReportNavState("Before FireEvent");
                            hsec.FireEvent("onchange", ref eventObj);
                            awc.ReportNavState("After FireEvent");
                            return true;
                            }
                        }
                    }
                }
            awc.ReportNavState("never found control for set");
            return fNeedSave;
        }
#if no
		private bool FClickControlName(IHTMLDocument2 oDoc2, string sTag, string sName)
		{
			// find an sTag with name sName
			IHTMLElementCollection hec;
			hec = (IHTMLElementCollection)oDoc2.all.tags(sTag);

			if (sTag.ToUpper() == "a")
				{
				foreach (IHTMLAnchorElement 
				}
		}

#endif // no
        /* F  C L I C K  C O N T R O L */
        /*----------------------------------------------------------------------------
        	%%Function: FClickControl
        	%%Qualified: ArbWeb.WebControl.FClickControl
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public bool FClickControl(IHTMLDocument2 oDoc2, string sId, string sidWaitFor = null)
        {
            m_srpt.LogData(String.Format("FClickControl {0}", sId), 3, StatusRpt.MSGT.Body);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return FWaitForNavFinish(sidWaitFor);
        }


        public bool FClickControlNoWait(IHTMLDocument2 oDoc2, string sId)
        {
            m_srpt.LogData(String.Format("FClickControlNoWait: {0}", sId), 3, StatusRpt.MSGT.Body);
            return FClickControlInDocNoWait(oDoc2, sId);
        }

        /* F  C L I C K  C O N T R O L  N O  W A I T */
        /*----------------------------------------------------------------------------
        	%%Function: FClickControlNoWait
        	%%Qualified: ArbWeb.WebControl.FClickControlNoWait
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        static public bool FClickControlInDocNoWait(IHTMLDocument2 oDoc2, string sId)
        {
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return true;
        }
        /* F  C H E C K  F O R  C O N T R O L */
        /*----------------------------------------------------------------------------
        	%%Function: FCheckForControl
        	%%Qualified: ArbWeb.WebControl.FCheckForControl
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static bool FCheckForControl(IHTMLDocument2 oDoc2, string sId)
        {
            if (oDoc2?.all.item(sId, 0) != null)
                return true;
                
            return false;
        }
        /* S  G E T  C O N T R O L  V A L U E */
        /*----------------------------------------------------------------------------
        	%%Function: SGetControlValue
        	%%Qualified: ArbWeb.WebControl.SGetControlValue
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string SGetControlValue(IHTMLDocument2 oDoc2, string sId)
        {
            if (FCheckForControl(oDoc2, sId))
                return (string)((IHTMLInputElement)oDoc2.all.item(sId, 0)).value;
            return null;
        }

        public static string SGetElementValue(IHTMLDocument2 oDoc2, string sId)
        {
            if (FCheckForControl(oDoc2, sId))
                return (string)((IHTMLElement)oDoc2.all.item(sId, 0)).innerText;
            return null;
        }

        public static IHTMLElement ElementGet(IHTMLDocument2 oDoc2, string sId)
        {
            if (FCheckForControl(oDoc2, sId))
                return (IHTMLElement)oDoc2.all.item(sId, 0);

            return null;
        }

        public static string SGetSelectSelectedValue(IHTMLDocument2 oDoc2, string sId)
        {
            if (FCheckForControl(oDoc2, sId))
                return (string)((HTMLSelectElementClass)oDoc2.all.item(sId, 0)).value;
            return null;
        }


        /* R G S  F R O M  C H L B X */
        /*----------------------------------------------------------------------------
        	%%Function: RgsFromChlbx
        	%%Qualified: ArbWeb.WebControl.RgsFromChlbx
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string[] RgsFromChlbx(bool fUse, CheckedListBox chlbx)
        {
            return RgsFromChlbx(fUse, chlbx, -1, false, null, false);
        }
        /* R G S  F R O M  C H L B X  S P O R T */
        /*----------------------------------------------------------------------------
        	%%Function: RgsFromChlbxSport
        	%%Qualified: ArbWeb.WebControl.RgsFromChlbxSport
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string[] RgsFromChlbxSport(bool fUse, CheckedListBox chlbx, string sSport, bool fMatch)
        {
            return RgsFromChlbx(fUse, chlbx, -1, false, sSport, fMatch);
        }
        /* R G S  F R O M  C H L B X */
        /*----------------------------------------------------------------------------
        	%%Function: RgsFromChlbx
        	%%Qualified: ArbWeb.WebControl.RgsFromChlbx
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string[] RgsFromChlbx(
            bool fUse,
            CheckedListBox chlbx,
            int iForceToggle,
            bool fForceOn,
            string sSport,
            bool fMatch)
        {
            string sSport2 = sSport == "Softball" ? "SB" : sSport;

            if (!fUse && sSport == null)
                return null;

            int c = chlbx.CheckedItems.Count;

            if (!fUse)
                c = chlbx.Items.Count;

            if (iForceToggle != -1)
                {
                if (fForceOn)
                    c++;
                else
                    c--;
                }

            string[] rgs = new string[c];
            int i = 0;

            if (!fUse)
                {
                int iT = 0;

                for (i = 0; i < c; i++)
                    {
                    rgs[iT] = (string) chlbx.Items[i];
                    if (sSport != null)
                        {
                        if ((rgs[iT].IndexOf(sSport) >= 0 && fMatch)
                            || (rgs[iT].IndexOf(sSport) == -1 && !fMatch)
                            || (rgs[iT].IndexOf(sSport2) >= 0 && fMatch)
                            || (rgs[iT].IndexOf(sSport2) == -1 && !fMatch))
                            {
                            iT++;
                            }
                        }
                    else
                        {
                        iT++;
                        }
                    }
                if (iT < c)
                    Array.Resize(ref rgs, iT);

                return rgs;
                }

            i = 0;
            foreach (int iChecked in chlbx.CheckedIndices)
                {
                if (iChecked == iForceToggle)
                    continue;
                rgs[i] = (string) chlbx.Items[iChecked];
                if (sSport != null)
                    {
                    if ((rgs[i].IndexOf(sSport) >= 0 && fMatch)
                        || (rgs[i].IndexOf(sSport) == -1 && !fMatch))
                        {
                        i++;
                        }
                    }
                else
                    {
                    i++;
                    }
                }
            if (fForceOn && iForceToggle != -1)
                rgs[i++] = (string) chlbx.Items[iForceToggle];

            if (i < c)
                Array.Resize(ref rgs, i);

            return rgs;
        }

        internal class Utils
        {
            /* P L S  U N I Q U E  F R O M  R G S */
            /*----------------------------------------------------------------------------
                        %%Function: PlsUniqueFromRgs
                        %%Qualified: ArbWeb.CountsData:GameData:Games.PlsUniqueFromRgs
                        %%Contact: rlittle

            ----------------------------------------------------------------------------*/
            public static SortedList<string, int> PlsUniqueFromRgs(string[] rgs)
            {
                if (rgs == null)
                    return null;

                SortedList<string, int> pls = new SortedList<string, int>();
                foreach (string s in rgs)
                {
                    if (!pls.ContainsKey(s))
                        pls.Add(s, 0);
                }
                return pls;
            }
        }
        /* U P D A T E  C H L B X  F R O M  R G S */
        /*----------------------------------------------------------------------------
        	%%Function: UpdateChlbxFromRgs
        	%%Qualified: ArbWeb.WebControl.UpdateChlbxFromRgs
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static void UpdateChlbxFromRgs(CheckedListBox chlbx, string[] rgsSource, string[] rgsChecked, string[] rgsFilterPrefix, bool fCheckAll)
        {
            chlbx.Items.Clear();
            SortedList<string, int> mp = Utils.PlsUniqueFromRgs(rgsChecked);
			
            foreach (string s in rgsSource)
                {
                bool fSkip = false;

                if (rgsFilterPrefix != null)
                    {
                    fSkip = true;
                    foreach (string sPrefix in rgsFilterPrefix)
                        {
                        if (s.Length > sPrefix.Length && String.Compare(s.Substring(0, sPrefix.Length), sPrefix, true/*ignoreCase*/) == 0)
                            {
                            fSkip = false;
                            break;
                            }
                        }
                    }
                if (fSkip)
                    continue;

                CheckState cs;

                if (fCheckAll || mp.ContainsKey(s))
                    cs = CheckState.Checked;
                else
                    cs = CheckState.Unchecked;

                int i = chlbx.Items.Add(s, cs);
                }
        }

        /* O N  N A V I G A T I N G */
        /*----------------------------------------------------------------------------
        	%%Function: OnNavigating
        	%%Qualified: ArbWeb.WebControl.OnNavigating
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void OnNavigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            m_srpt.AddMessage(String.Format("OnNavigating: {0}", e.Url));

        }

        /* O N  N A V I G A T E D */
        /*----------------------------------------------------------------------------
        	%%Function: OnNavigated
        	%%Qualified: ArbWeb.WebControl.OnNavigated
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void OnNavigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            m_srpt.AddMessage(String.Format("OnNavigated: {0}", e.Url));
        }

        /* O N  N E W  W I N D O W */
        /*----------------------------------------------------------------------------
        	%%Function: OnNewWindow
        	%%Qualified: ArbWeb.WebControl.OnNewWindow
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void OnNewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
           m_srpt.AddMessage(String.Format("OnNewWindow called"));
        }

        /* O N  D O W N L O A D */
        /*----------------------------------------------------------------------------
        	%%Function: OnDownload
        	%%Qualified: ArbWeb.WebControl.OnDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void OnDownload(object sender, EventArgs e)
        {
            m_srpt.AddMessage(String.Format("OnDownload called"));
        }

    }


}
