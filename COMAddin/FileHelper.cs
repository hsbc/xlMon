
using System.IO;

namespace XLMonCOMAddin
{
    class FileHelpers
    {
        public static string GetAssemblyFolder()
        {
            return Path.GetDirectoryName(GetAssemblyFullPath());
        }

        public static string GetAssemblyFullPath()
        {
            //The functions in the .net Framework don't handle this well, returning different formats depending
            //on whether we are running locally or from a network share, hence we code this ourselves.

            string fullPathOfAssembly = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            if (fullPathOfAssembly.StartsWith("file:///")) fullPathOfAssembly = fullPathOfAssembly.Substring(8);
            if (fullPathOfAssembly.StartsWith("file:")) fullPathOfAssembly = fullPathOfAssembly.Substring(5);

            return fullPathOfAssembly;
        }
    }
}
