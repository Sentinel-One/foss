using System;
using System.Threading;
using System.Windows.Forms;

namespace S1ExcelPlugIn
{
    class QueryThread
    {
        ContainerControl m_sender = null;
        Delegate m_senderDelegate = null;

        public QueryThread()
        {
        }
        public QueryThread(ContainerControl sender, Delegate senderDelegate)
        {
            m_sender = sender;
            m_senderDelegate = senderDelegate;
        }

        public QueryThread(ContainerControl sender, Delegate senderDelegate, params object[] list)
        {
            m_sender = sender;
            m_senderDelegate = senderDelegate;
        }

        public void RunProcess(object obj)
        {
            Thread.CurrentThread.IsBackground = true; //make them a daemon
            object[] objArray = (object[])obj;
            m_sender = (System.Windows.Forms.Form)objArray[0];
            m_senderDelegate = (System.Delegate)objArray[1];

            LocalRunProcess();
        }

        public void RunProcess()
        {
            Thread.CurrentThread.IsBackground = true; //make them a daemon
            LocalRunProcess();
        }

        private void LocalRunProcess()
        {
            try
            {
                m_sender.BeginInvoke(m_senderDelegate);
            }
            catch (System.InvalidOperationException)
            {
                Thread.CurrentThread.Abort();
            }
            catch
            {
                Thread.CurrentThread.Abort();
            }
        }
    }
}
