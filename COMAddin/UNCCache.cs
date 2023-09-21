using System.Collections.Generic;
using System.IO;

namespace XLMonCOMAddin
{
    class UNCCache
    {
        readonly Dictionary<string, string> lookup;

        public UNCCache()
        {
            lookup = new Dictionary<string, string>();
        }

        public string GetUNC(string originalPath)
        {
            if (!lookup.ContainsKey(originalPath))
            {
                try
                {   //This is in a try block because there are some types of filepath that can't 
                    //be used as a parameter to FileInfo, such as when opening from confluence or
                    //sharepoint the path begins with http://
                    FileInfo fi = new FileInfo(originalPath);
                    lookup[originalPath] = fi.UncPath();
                }
                catch
                {
                    //If we couldn't generate a UNC name for whatever reason then just use the one we have been given
                    lookup[originalPath] = originalPath;
                }
            }
            return lookup[originalPath];
        }
    }
}
