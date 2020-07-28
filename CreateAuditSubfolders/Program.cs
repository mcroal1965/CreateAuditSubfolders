using System;
using System.IO;
using System.Diagnostics;


namespace CreateAuditSubfolders
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string Input = args[0];
                if (!Directory.Exists(Input))
                {
                    Console.WriteLine(@"Folder does not exist: "+Input);
                    Console.WriteLine(@"Syntax: CreateAuditSubFolders drive:\path");
                    Console.WriteLine("");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
                Int32 InputLength = Input.Length;

                //generate CMD file that uses nconvert to change TIFs to PDFs
                String[] FoundTIFs = Directory.GetFiles(Input, "*.tif", SearchOption.TopDirectoryOnly);
                String TIFPDFcmd = Input + @"\TifToPDF.cmd";
                File.Delete(TIFPDFcmd);
                File.AppendAllText(TIFPDFcmd, "c:\r\n");
                File.AppendAllText(TIFPDFcmd, @"cd\cb\nConvert"+"\r\n");
                foreach (String TIFFile in FoundTIFs)
                {
                    File.AppendAllText(TIFPDFcmd, "nconvert -xall -multi -o $%% -out pdf -D -c 3 " + TIFFile+"\r\n");
                }
                File.AppendAllText(TIFPDFcmd, "exit");

                //run the CMD file as administrator
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.WorkingDirectory = @"C:\Windows\System32";
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/user:Administrator \"cmd /K " + TIFPDFcmd + "\"";
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();

                String[] FoundFiles = Directory.GetFiles(Input, "*.*", SearchOption.TopDirectoryOnly);
                foreach (String FoundFile in FoundFiles)
                {
                    string LoanNumber = FoundFile.Substring(InputLength + 1, FoundFile.IndexOf("-") - InputLength - 1);
                    string NewDirectory = Input + @"\" + LoanNumber;
                    string FileName = FoundFile.Substring(InputLength);

                    if (!Directory.Exists(NewDirectory))
                        {
                            Directory.CreateDirectory(NewDirectory);
                        }
                    File.Copy(FoundFile, NewDirectory + @"\" + FileName);
                    File.Delete(FoundFile);
                }

                Environment.Exit(0);
            }
            catch(System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(1);
            }
        }
    }
}
