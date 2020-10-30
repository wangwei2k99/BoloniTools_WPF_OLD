using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BoloniTools.Controller
{
    public static class Notice
    {
        public static event EventHandler<string> NoticeEvent;
        public static void NoticeFunc(string str)
        {
            if (NoticeEvent!=null)
            {
                NoticeEvent(null, str);
            }
        }
        public static void NoticeDemo()
        {
            NoticeFunc("测试通知");
        }
    }
}
