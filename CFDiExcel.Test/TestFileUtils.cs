using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace CFDiExcel.Test
{
    public static class TestFileUtils
    {
        public static string ReadFileAsString(string file, [CallerFilePath] string filePath = "")
        {
            var directoryPath = Path.GetDirectoryName(filePath);
            var fullPath = Path.Join(directoryPath, "EjemplosCFDi", file);
            return File.ReadAllText(fullPath);
        }
    }
}
