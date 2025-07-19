using Eplan.EplApi.ApplicationFramework;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Eplan.EplAddin.ERP
{
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }

        public bool OnInit()
        {
            // Ścieżka do folderu z DLL (zmień jeśli używasz innego katalogu)
            string dllFolder = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Addins", "Eplan.EplAddin.ERP", "Libs");

            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    string assemblyName = new AssemblyName(args.Name).Name + ".dll";
                    string dllPath = Path.Combine(dllFolder, assemblyName);

                    if (File.Exists(dllPath))
                    {
                        return Assembly.LoadFrom(dllPath);
                    }
                    else
                    {
                        LogMissingAssembly(assemblyName, dllFolder, "File not found");
                    }
                }
                catch (Exception ex)
                {
                    LogMissingAssembly(args.Name, dllFolder, ex.ToString());
                }

                return null;
            };

            // 💡 Debug info: pokaż z jakiej lokalizacji ładowany jest ClosedXML
            try
            {
                var loadedPath = typeof(ClosedXML.Excel.XLWorkbook).Assembly.Location;
                MessageBox.Show("ClosedXML załadowano z:\n" + loadedPath, "Diagnostyka DLL");
            }
            catch (Exception ex)
            {
                MessageBox.Show("ClosedXML NIE załadowano.\n" + ex.ToString(), "Błąd DLL");
            }

            return true;
        }

        public bool OnInitGui()
        {
            new CommandLineInterpreter().Execute("RegisterAction /Name:CheckAvailability /Namespace:Eplan.EplAddin.ERP.CheckAvailability");
            new CommandLineInterpreter().Execute("RegisterAction /Name:AnalyzeBomWithPrices /Namespace:Eplan.EplAddin.ERP.AnalyzeBomWithPrices");
            return true;
        }

        public bool OnExit()
        {
            return true;
        }

        private void LogMissingAssembly(string name, string path, string reason)
        {
            try
            {
                string logPath = @"C:\EplanData\addin_errors.log";
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                using (var sw = new StreamWriter(logPath, true))
                {
                    sw.WriteLine($"[{DateTime.Now}] Błąd ładowania DLL: {name}");
                    sw.WriteLine($"  Szukano w: {path}");
                    sw.WriteLine($"  Powód: {reason}");
                    sw.WriteLine();
                }
            }
            catch
            {
                // Nie logujemy błędów z logowania ;)
            }
        }
    }
}
