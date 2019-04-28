using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumPOMframework.Common
{
    class ReadGlobalProperties
    {
        public String getGlobalProperty(String prop)
        {
            ResourceManager rm = new ResourceManager("SeleniumPOMframework.properties", Assembly.GetExecutingAssembly());

            // Retrieve the value of the string resource named "welcome".
            // The resource manager will retrieve the value of the  
            // localized resource using the caller's current culture setting.
            String url = rm.GetString(prop);
            return url;
        }
    }
}
