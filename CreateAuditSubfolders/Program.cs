using System;
using System.IO;
using System.Diagnostics;


namespace CreateAuditSubfolders
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine(@"Incorrect number of arguments supplied.");
                Console.WriteLine(@"Syntax: CreateAuditSubFolders F:\NautilusExports\{auditX} YN");
                Console.WriteLine("");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(1);
            }
            try
            {
                string Input = args[0];
                if (!Directory.Exists(Input))
                {
                    Console.WriteLine(@"Folder does not exist: "+Input);
                    Console.WriteLine(@"Syntax: CreateAuditSubFolders F:\NautilusExports\{auditX} YN");
                    Console.WriteLine("");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
                Int32 InputLength = Input.Length;

                string ConvertYN = args[1];
                if (ConvertYN != "Y" && ConvertYN != "N")
                {
                    Console.WriteLine(@"ConvertYN argument is not Y or N: " + ConvertYN);
                    Console.WriteLine(@"Syntax: CreateAuditSubFolders F:\NautilusExports\{auditX} YN");
                    Console.WriteLine("");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }

                if (!Directory.Exists(@"F:\NautilusExports\nConvert"))
                {
                    Console.WriteLine(@"nConvert program folder does not exist at F:\NautilusExports\nConvert");
                    Console.WriteLine("");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }

                //rename .___ to .TIF
                String[] FoundFFFs = Directory.GetFiles(Input, "*.___", SearchOption.TopDirectoryOnly);
                foreach (String FFFile in FoundFFFs)
                {
                    FileInfo fi = new FileInfo(FFFile);
                    String src = fi.FullName;
                    String dest = fi.FullName.Substring(0,fi.FullName.Length-fi.Extension.Length)+".tif";
                    File.Copy(src, dest);
                    File.Delete(src);
                }

                //remove periods in filename
                String[] FoundFPs = Directory.GetFiles(Input, "*.*.*", SearchOption.TopDirectoryOnly);
                foreach (String FPFile in FoundFPs)
                {
                    FileInfo fi = new FileInfo(FPFile);
                    String fiext = fi.Extension;
                    String finame = fi.FullName.Substring(0, fi.FullName.Length - fi.Extension.Length);
                    if (finame.Contains("."))
                        {
                            String src = fi.FullName;
                            String dest = fi.FullName.Substring(0, fi.FullName.Length - fi.Extension.Length).Replace(".","") + fi.Extension;
                            File.Copy(src, dest);
                            File.Delete(src);
                        }
                }

                //generate CMD file that uses nconvert to change TIFs to PDFs
                if (ConvertYN == "Y")
                {
                    String[] FoundTIFs = Directory.GetFiles(Input, "*.tif", SearchOption.TopDirectoryOnly);
                    String TIFPDFcmd = Input + @"\TifToPDF.cmd";
                    File.Delete(TIFPDFcmd);
                    File.AppendAllText(TIFPDFcmd, "F:\r\n");
                    File.AppendAllText(TIFPDFcmd, @"cd\NautilusExports\nConvert" + "\r\n");
                    foreach (String TIFFile in FoundTIFs)
                    {
                        long size = new System.IO.FileInfo(TIFFile).Length;
                        if (size > 2048000)
                        {
                            File.AppendAllText(TIFPDFcmd, "nconvert -xall -multi -o $%% -out pdf -D -c 3 " + TIFFile + "\r\n");
                        }
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

                    File.Delete(TIFPDFcmd);
                }
                

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
