using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Xml.Linq;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace VBUC_ReferenceSolver
{

    class Program
    {
        const string TLB_REPO_URL = "https://github.com/orellabac/TLBRepo/archive/master.zip";
        static void Usage()
        {
            Console.WriteLine("VBUC Reference Solver for VBUC");
            Console.WriteLine("==============================");
            Console.WriteLine("By Mauricio Rojas");

            Console.WriteLine("This tools download all tlbs from a TLB Repo and tries to resolved all COM references with them");
            Console.WriteLine("Usage: VBUC_ReferenceSolver.exe <solution.VBUCSln>");
        }

        public class TypeLibInfo
        {
            public string File { get; set; }
            public string GUID { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }

        }
        private enum RegKind
        {
            RegKind_Default = 0,
            RegKind_Register = 1,
            RegKind_None = 2
        }
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(String strTypeLibName, RegKind regKind,
           [MarshalAs(UnmanagedType.Interface)] out ITypeLib typeLib);
        static void Main(string[] args)
        {
            Usage();
            if (args.Length != 1)
            {
                Console.WriteLine("Invalid args");
                return;
            }
            // Download repo
            var repoZIP = "TLB_REPO.zip";
            var outputRepoPath = ".";
            if (!File.Exists(repoZIP))
            {
                Console.WriteLine("Downloading repo...");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                using (var client = new WebClient())
                {
                    client.DownloadFile(TLB_REPO_URL, repoZIP);
                }
                Console.WriteLine("Extracting ...");
                using (ZipFile zip = ZipFile.Read(repoZIP))
                {
                    zip.ExtractAll(outputRepoPath);
                }
            }
            Console.WriteLine("Loading TLB infos");
            var tlbs = new List<TypeLibInfo>();
            foreach (var file in Directory.GetFiles(outputRepoPath, "*.tlb", SearchOption.AllDirectories))
            {

                ITypeLib typeLib = null;

                try
                {
                    LoadTypeLibEx(file, RegKind.RegKind_Default, out typeLib);
                    IntPtr ppTLibAttr;
                    typeLib.GetLibAttr(out ppTLibAttr);
                    TYPELIBATTR tlibattr = (TYPELIBATTR)Marshal.PtrToStructure(ppTLibAttr, typeof(TYPELIBATTR));
                    typeLib.ReleaseTLibAttr(ppTLibAttr);
                    tlbs.Add(
                   new TypeLibInfo()
                   {
                       File = file,
                       GUID = tlibattr.guid.ToString("B").ToUpper(),
                       Major = tlibattr.wMajorVerNum,
                       Minor = tlibattr.wMinorVerNum
                   });
                }
                catch
                { }
                finally
                {
                    if (typeLib != null)
                    {
                        Marshal.ReleaseComObject(typeLib);
                    }
                }

            }
            var inputFile = args[0];
            XDocument solution;
            using (var f = System.IO.File.OpenRead(inputFile))
            {
                solution = XDocument.Load(f);
            }
            foreach (var reference in solution.Element("VBUpgradeSolution").Elements("ExternalReference"))
            {
                var guid = reference.Attribute("Guid").Value.ToUpper(); 
                int.TryParse(reference.Attribute("MinorVersion").Value, out var minor);
                int.TryParse(reference.Attribute("MajorVersion").Value, out var major);
                var refInTlbs = tlbs.Find(x => (x.GUID == guid && x.Minor == minor && x.Major == major));
                if (refInTlbs != null)
                {
                    Console.WriteLine($"Mapping {reference.Attribute("FriendlyName")} using TLB {refInTlbs.File} ");
                    reference.Attribute("AbsolutePath").Value = Path.GetFullPath(refInTlbs.File);
                }
            }
            solution.Save(inputFile);
            Console.WriteLine("Solution file updated");
        }
    }
}

