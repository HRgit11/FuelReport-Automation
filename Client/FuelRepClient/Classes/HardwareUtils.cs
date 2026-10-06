using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace mamaRabota.Classes
{
    public static class HardwareUtils
    {
        public static string GetHwId()
        {
            try
            {
                string cpuInfo = "";
                using (ManagementClass mc = new ManagementClass("Win32_Processor"))
                {
                    using (ManagementObjectCollection moc = mc.GetInstances())
                    {
                        foreach (ManagementObject mo in moc)
                        {
                            cpuInfo = mo.Properties["ProcessorId"].Value.ToString();
                            break;
                        }
                    }
                }
                return cpuInfo;
            }
            catch
            {
                return "ошибка получения HwId";
            }

        }
    }
}
