using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;


namespace StatusBox
{
    public class StatusRpt
    {
        public enum MSGT
        {
            Header,
            Body,
            Warning,
            Error
        };

        RichTextBox m_reReport;
        int m_iLevel = 0;
        bool m_fInit = false;
        private string m_sFile;
        private int m_nLogLevel = 3;
        private MSGT m_msgtFilter; // Error = Errors; Warning = Error + Warnings; Header = Warning + Header; Body = Header + Body

        public void SetLogLevel(int nLevel)
        {
            m_nLogLevel = nLevel;
        }

        public void SetFilter(MSGT msgt)
        {
            m_msgtFilter = msgt;
        }

        public void AddMessage(string s, MSGT msgt, bool fFileOnly = false)
        {
            if (!m_fInit)
                return;
            if (m_reReport == null)
                return;
            AddMessageCore(s, msgt, m_iLevel, fFileOnly);
        }


        private bool FShouldLog(int nLogLevel, MSGT msgtSev, out bool fFileOnly)
        {
            fFileOnly = true;

            if (nLogLevel > m_nLogLevel)
                return false;

            if ((nLogLevel % 1) == 1)
                fFileOnly = false;

            switch (m_msgtFilter)
            {
                case MSGT.Body:
                    if (msgtSev == MSGT.Body)
                        return true;
                    goto case MSGT.Header;
                case MSGT.Header:
                    if (msgtSev == MSGT.Header)
                        return true;
                    goto case MSGT.Warning;
                case MSGT.Warning:
                    if (msgtSev == MSGT.Warning)
                        return true;
                    goto case MSGT.Error;
                case MSGT.Error:
                    if (msgtSev == MSGT.Error)
                        return true;
                    break;
            }
            return false;
        }

        public void LogData(string sLog, int nLogLevel, MSGT msgtSev)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
        }

        public void LogData(string sLog, int nLogLevel, MSGT msgtSev, IDictionary<string, string> mps)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
            PushLevel();
            foreach (string sKey in mps.Keys)
                AddMessage(String.Format("{{ \"{0}\"->\"{1}\" }},", sKey, mps[sKey]), MSGT.Body, fFileOnly);
            PopLevel();
        }

        public void LogData(string sLog, int nLogLevel, MSGT msgtSev, IDictionary<string, int> mps)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
            PushLevel();
            foreach (string sKey in mps.Keys)
                AddMessage(String.Format("{{ \"{0}\"->\"{1}\" }},", sKey, mps[sKey]), MSGT.Body, fFileOnly);
            PopLevel();
        }
        public void LogData(string sLog, int nLogLevel, MSGT msgtSev, IDictionary<int, Dictionary<string, string>> mpn)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
            PushLevel();
            foreach (int nKey in mpn.Keys)
            {
                Dictionary<string, string> mps = mpn[nKey];
                AddMessage("Dictionary<string,string>:", msgtSev, fFileOnly);
                PushLevel();
                foreach (string sKey in mps.Keys)
                {
                    AddMessage(String.Format("{{ \"{0}\"->\"{1}\" }},", sKey, mps[sKey]), MSGT.Body, fFileOnly);
                }
            }
            PopLevel();
        }

        public void LogData(string sLog, int nLogLevel, MSGT msgtSev, IDictionary<int, List<string>> mpn)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
            PushLevel();
            foreach (int nKey in mpn.Keys)
            {
                List<string> pls = mpn[nKey];
                AddMessage("List<string>:", msgtSev, fFileOnly);
                PushLevel();
                foreach (string s in pls)
                {
                    AddMessage(String.Format("{{\"{0}\" }},", s), MSGT.Body, fFileOnly);
                }
            }
            PopLevel();
        }

        public void LogData(string sLog, int nLogLevel, MSGT msgtSev, IEnumerable<string> pls)
        {
            bool fFileOnly;
            if (!FShouldLog(nLogLevel, msgtSev, out fFileOnly))
                return;

            AddMessage(sLog, msgtSev, fFileOnly);
            PushLevel();
            foreach (string s in pls)
                AddMessage(String.Format("{{\"{0}\" }},", s), MSGT.Body, fFileOnly);
            PopLevel();
        }

        public void AddMessage(string s, bool fFileOnly = false)
        {
            if (!m_fInit)
                return;
            if (m_reReport == null)
                return;
            AddMessageCore(s, MSGT.Body, m_iLevel, fFileOnly);
        }

        public void AddMessage(string s, MSGT msgt, int iLevel, bool fFileOnly = false)
        {
            if (!m_fInit)
                return;
            if (m_reReport == null)
                return;
            AddMessageCore(s, msgt, iLevel, fFileOnly);
        }

        public delegate void AddMessageDelegate(string s, MSGT msgt, int iLevel);

        private void AddLogFileCore(string s)
        {
            StreamWriter sw = new StreamWriter(m_sFile, true /*fAppend*/, System.Text.Encoding.Default);
            sw.Write(s);
            sw.Close();
        }
        private Object oLogLock = new Object();

        private void AddLogfile(string s, MSGT msgt, int iLevel)
        {
            string sTab = "";
            string sText;

            for (int cTab = 0; cTab < iLevel; cTab++)
                sTab += "\t";

            switch (msgt)
            {
                default:
                    sText = sTab + s + "\n";
                    break;
                case MSGT.Error:
                    sText = sTab + "ERROR : " + s + "\n";
                    break;
                case MSGT.Warning:
                    sText = sTab + "WARNING : " + s + "\n";
                    break;
            }

            lock (oLogLock)
            {
                AddLogFileCore(sText);
            }
        }

        private void AddMessageCore(string s, MSGT msgt, int iLevel, bool fFileOnly)
        {
            if (m_sFile != null)
                AddLogfile(s, msgt, iLevel);

            if (fFileOnly)
                return;

            if (!m_fInit)
                return;
            if (m_reReport == null)
                return;

            object[] rgo = new object[3];

            rgo[0] = s;
            rgo[1] = msgt;
            rgo[2] = iLevel;

            m_reReport.BeginInvoke(new AddMessageDelegate(DoAddMessage), rgo);
        }

        private void DoAddMessage(string s, MSGT msgt, int iLevel)
        {
            if (!m_fInit)
                return;

            //                if (m_reReport.Text.Length > 0)
            //                   m_reReport.AppendText("\n");

            if (iLevel == -1)
                iLevel = m_iLevel;

            string sTab = "";
            string sTabStops = "\\tx360\\tx720\\tx1080\\tx1440\\tx1760 ";
            string sHeaderFmt = "\\b\\fs20";
            string sBodyFmt = "\\fs18";
            string sColorTbl = "{\\colortbl;\\red255\\green0\\blue0;\\red255\\green128\\blue0;}";
            string sErrorFmt = "\\b\\cf1 ";
            string sWarningFmt = "\\b\\cf2 ";
            for (int cTab = 0; cTab < iLevel; cTab++)
                sTab += "\\tab ";

            switch (msgt)
            {
                case MSGT.Header:
                    m_reReport.SelectedRtf = "{\\rtf1" + sColorTbl + "\\pard" + sHeaderFmt + sTabStops + sTab + s + "\\par}";
                    iLevel++;
                    break;
                case MSGT.Body:
                    m_reReport.SelectedRtf = "{\\rtf1" + sColorTbl + "\\pard" + sBodyFmt + sTabStops + sTab + s + "\\par}";
                    break;
                case MSGT.Error:
                    m_reReport.SelectedRtf = "{\\rtf1" + sColorTbl + "\\pard" + sBodyFmt + sErrorFmt + sTabStops + sTab + s + "\\par}";
                    break;
                case MSGT.Warning:
                    m_reReport.SelectedRtf = "{\\rtf1" + sColorTbl + "\\pard" + sBodyFmt + sWarningFmt + sTabStops + sTab + s + "\\par}";
                    break;
            }
            m_reReport.ScrollToCaret();
            //		    if (msgt == MSGT.Header)
            //              m_reReport.Refresh();
        }

        public void PushLevel()
        {
            m_iLevel++;
        }

        public void PopLevel()
        {
            m_iLevel--;
            if (m_iLevel < 0)
                m_iLevel = 0;
        }

        public StatusRpt()
        {
            m_reReport = null;
            m_iLevel = 0;
            m_fInit = false;
        }

        public void Init(RichTextBox re)
        {
            m_reReport = re;
            m_iLevel = 0;
            if (re != null)
                m_fInit = true;
        }

        public StatusRpt(RichTextBox re)
        {
            m_reReport = re;
            var x = re.Handle;

            m_iLevel = 0;
            if (re != null)
                m_fInit = true;
        }

        public void AttachLogfile(string sFile)
        {
            m_sFile = sFile;
        }
        public void UnitTest()
        {
            AddMessage("Generic message add");
            AddMessage("2nd line add");
            AddMessage("2nd line add");
            AddMessage("Heading add, no level", MSGT.Header);
            AddMessage("Under heading, no level");
            AddMessage("Under heading, error", MSGT.Error);
            AddMessage("Under heading, warning", MSGT.Warning);
            PushLevel();
            AddMessage("Heading add (level pushed)", MSGT.Header);
            AddMessage("Under heading");
            AddMessage("Under heading, no level");
            AddMessage("Under heading, error", MSGT.Error);
            AddMessage("Under heading, warning", MSGT.Warning);
            PopLevel();
            AddMessage("Heading add (level popped)", MSGT.Header);
            AddMessage("Under heading");
            AddMessage("Under heading, no level");
            AddMessage("Under heading, error", MSGT.Error);
            AddMessage("Under heading, warning", MSGT.Warning);
        }
    }
}

