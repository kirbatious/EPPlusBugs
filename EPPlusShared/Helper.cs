using System.IO;
using System;

namespace EPPlusShared
{
    public static class Helper
    {
        public static FileInfo GetFileInfo(string fileName)
        {
            var dir = typeof(Helper).Assembly.Location;
            var index = dir.IndexOf(@"\bin\", StringComparison.OrdinalIgnoreCase);
            var baseDir = Directory.GetParent(dir[..index]);
            var path = Path.Combine(baseDir!.FullName, "EPPlusShared", fileName);
            var fi = new FileInfo(path);
            return fi;
        }

    }
}
