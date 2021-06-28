using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;


namespace UseCom
{

    /// <summary>
    /// Managed definition of COM interface
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("fa54fb37-fb48-4fef-b73e-c91add16f718")]
    internal interface IO1
    {
        void StoreSomething(string s);
    }

    [ComImport]
    [CoClass(typeof(O1Class))]
    [Guid("fa54fb37-fb48-4fef-b73e-c91add16f718")]
    internal interface O1 : IO1
    {
    }


    [ComImport]
    [Guid("ca496a65-7caa-4f3c-8810-459661bb6afa")]
    internal class O1Class
    {
    }


    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var io = (O1)new O1Class();
                io.StoreSomething("blah");
            }
            catch (Exception ex)
            {
                var fullname = System.Reflection.Assembly.GetEntryAssembly().Location;
                var progname = Path.GetFileNameWithoutExtension(fullname);
                Console.Error.WriteLine($"{progname} Error: {ex.Message}");
            }

        }
    }
}
