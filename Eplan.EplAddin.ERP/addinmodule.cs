using Eplan.EplApi.ApplicationFramework;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Eplan.EplAddin.ERP
{
    public class AddInModule : IEplAddIn
    {
        // Ten katalog będzie tym samym, w którym wrzucisz wszystkie DLL (wraz z ERP.dll)
        private static readonly string _addinsDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister() => true;

        public bool OnInit()
        {
            // Podpinamy resolver, który będzie szukał zależności w katalogu dodatku erp
            AppDomain.CurrentDomain.AssemblyResolve += Resolver;
            return true;
        }

        private Assembly Resolver(object sender, ResolveEventArgs args)
        {
            // Przykład args.Name => "ClosedXML, Version=0.95.4.0, ..."

            string shortName = new AssemblyName(args.Name).Name + ".dll";
            string probe = Path.Combine(_addinsDir, shortName);

            if (File.Exists(probe))
            {
                try
                {
                    return Assembly.LoadFrom(probe);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Błąd ładowania {shortName}:\n{ex.Message}",
                                    "Resolver Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // jeśli go tu nie ma, zwracamy null i .NET pójdzie dalej (GAC / inne ścieżki)
            return null;
        }

        public bool OnInitGui()
        {
            var cli = new CommandLineInterpreter();
            cli.Execute("RegisterAction /Name:CheckAvailability /Namespace:Eplan.EplAddin.ERP.CheckAvailability");
            cli.Execute("RegisterAction /Name:AnalyzeBomWithPrices /Namespace:Eplan.EplAddin.ERP.AnalyzeBomWithPrices");
            return true;
        }

        public bool OnExit() => true;
    }
}
//todo Dodać obsługę deadline oraz budżetu