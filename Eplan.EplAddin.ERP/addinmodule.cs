using Eplan.EplApi.ApplicationFramework;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Eplan.EplAddin.ERP
{
    public class AddInModule : IEplAddIn
    {
        // katalog z DLL-ami (ERP.dll + wszystkie potrzebne biblioteki)
        private static readonly string _addinsDir =
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister() => true;

        public bool OnInit()
        {
            // rejestrujemy resolver, by ładował ClosedXML i inne z folderu dodatku
            AppDomain.CurrentDomain.AssemblyResolve += Resolver;
            return true;
        }

        private Assembly Resolver(object sender, ResolveEventArgs args)
        {
            string shortName = new AssemblyName(args.Name).Name + ".dll";
            string probe = Path.Combine(_addinsDir, shortName);
            if (File.Exists(probe))
            {
                try { return Assembly.LoadFrom(probe); }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Błąd ładowania {shortName}:\n{ex.Message}",
                        "Resolver Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return null;
        }

        public bool OnInitGui()
        {
            var cli = new CommandLineInterpreter();
            cli.Execute("RegisterAction /Name:CheckAvailability   /Namespace:Eplan.EplAddin.ERP.CheckAvailability");
            cli.Execute("RegisterAction /Name:AnalyzeBomWithPrices /Namespace:Eplan.EplAddin.ERP.AnalyzeBomWithPrices");
            return true;
        }

        public bool OnExit() => true;
    }
}
