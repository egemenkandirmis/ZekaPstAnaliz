using System;
using System.Collections.Generic;
using System.Text;

namespace MailBackupSystem
{
    class Class2
    {
        public double fuc(double num)
        {
            if (num == 1)
                return 1;

            return fuc(num-1) * num;
        }
    }
}
