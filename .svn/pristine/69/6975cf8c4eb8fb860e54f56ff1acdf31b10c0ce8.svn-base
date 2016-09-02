using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace Manual_Import.Helper
{
    class PdfTimeoutMethod
    {
        private ManualResetEvent MRE;
        Func<string, int[]> Fun_TidyPdf;
        int[] Back;
        public PdfTimeoutMethod(Func<string, int[]> func)
        {
            MRE = new ManualResetEvent(true);
            Fun_TidyPdf = func;
        }

        public int[] TidyWithTimeout(string path)
        {
            MRE.Reset();
            Fun_TidyPdf.BeginInvoke(path, DoAsyncCallBack, null);
            if (!MRE.WaitOne(new TimeSpan(0, 0, 20), false))
            {
                return new int[] {0,0 };
            }
            return Back;
        }

        private void DoAsyncCallBack(IAsyncResult result)
        {
            try
            {
                Back = Fun_TidyPdf.EndInvoke(result);
            }
            catch (Exception e)
            {
                string str=e.Message;
            }
            finally
            {
                MRE.Set();
            }

        }
    }
}
