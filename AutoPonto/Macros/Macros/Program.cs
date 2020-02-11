using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Macro
{
    class Test
    {
        static string matricula = "";
        static string senha = "";
        static string tipo = "";

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        static void Main(string[] args)
        {
            if (args.Length == 3)
            {
                foreach (string argument in args)
                {
                    string[] splitted = argument.Split('=');

                    if (splitted[0].ToLower() == "m" || splitted[0].ToLower() == "matricula")
                    {
                        matricula = splitted[1];
                    }
                    else if (splitted[0].ToLower() == "s" || splitted[0].ToLower() == "senha")
                    {
                        senha = splitted[1];
                    }
                    else if (splitted[0].ToLower() == "t" || splitted[0].ToLower() == "tipo")
                    {
                        tipo = splitted[1];
                    }
                }
            }
            else
            {
                return;
            }

            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Program Files (x86)\pw3270\pw3270.exe";
            proc.StartInfo.Arguments = " --host=mfv32701:23 --model=5";
            proc.Start();

            System.Threading.Thread.Sleep(1000);

            if (proc.MainWindowTitle == "pw3270 - mfv32701:23")
            {
                SetForegroundWindow(proc.MainWindowHandle);

                // escreve acics6
                SendKeys.SendWait("acics6~");
                System.Threading.Thread.Sleep(1000);

                // Escreve matrícula
                SendKeys.SendWait(String.Format("{0}", matricula));
                SendKeys.SendWait("{TAB}");
                System.Threading.Thread.Sleep(1000);

                // Escreve senha
                SendKeys.SendWait(String.Format("{0}", senha));
                SendKeys.SendWait("~");
                System.Threading.Thread.Sleep(1000);
                    
                // Seleciona recursos humanos
                SendKeys.SendWait("{DOWN}~");
                System.Threading.Thread.Sleep(1000);

                // Seleciona registro de ponto
                SendKeys.SendWait("{DOWN}~");
                System.Threading.Thread.Sleep(1000);

                // Seleciona o tipo (escreve e ou s)
                SendKeys.SendWait(String.Format("{0}", tipo));
                SendKeys.SendWait("~");
                System.Threading.Thread.Sleep(1000);
                    
                // Confirma
                SendKeys.SendWait("sim~");
                Console.WriteLine("confirmou");
                System.Threading.Thread.Sleep(1000);
                
                // Volta
                SendKeys.SendWait("{F3}");
                Console.WriteLine("voltou");
                System.Threading.Thread.Sleep(1000);

                // Sai
                SendKeys.SendWait("{DOWN}{DOWN}{DOWN}~");
                Console.WriteLine("saiu");
                System.Threading.Thread.Sleep(1000);
                
                proc.CloseMainWindow();
                proc.Close();

                return;
            }
        }
    }
}