using System;
using System.Reflection;
using System.Windows.Forms;

namespace ComComTag {
    static class Program {
        [STAThread]
        static void Main() {
            // Hook assembly resolver before any TagLibSharp types are JIT compiled
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            RunApp();
        }

        // Method separation ensures JIT compiler doesn't attempt to resolve TagLibSharp before hook is set
        static void RunApp() {
            Application.Run(new MainForm());
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args) {
            var executingAssembly = Assembly.GetExecutingAssembly();
            var assemblyName = new AssemblyName(args.Name);

            string[] possibleResourceNames = new[] {
                "ComComTag." + assemblyName.Name + ".dll",
                assemblyName.Name + ".dll",
                "ComComTag.TagLibSharp.dll",
                "TagLibSharp.dll"
            };

            foreach (string resName in possibleResourceNames) {
                using (var stream = executingAssembly.GetManifestResourceStream(resName)) {
                    if (stream != null) {
                        byte[] assemblyData = new byte[stream.Length];
                        stream.Read(assemblyData, 0, assemblyData.Length);
                        return Assembly.Load(assemblyData);
                    }
                }
            }

            return null;
        }
    }
}
