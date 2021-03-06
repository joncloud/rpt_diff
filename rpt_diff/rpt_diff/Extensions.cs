﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtensionMethods
{
    static class Extensions
    {   
        /*
         *  ToStringSafe
         *  - converts object to text representation
         *  - if object is null then returns empty string else returns converted object representation in string
         */
        public static string ToStringSafe(this Object obj)
        {
            return obj == null ? "" : obj.ToString();
        }
    }
}
