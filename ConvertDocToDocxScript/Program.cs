using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Security.AccessControl;
using System.Management.Automation;
using System.Configuration;

namespace ConvertDocToDocxScript
{
    class Program
    {
        static void Main(string[] args)
        {

            Microsoft.Office.Interop.Word.Application word = null;
            bool deleteFlag = true;
            bool copyACLFlag = false;
            var powerShell = PowerShell.Create();
            string pathToLogDir = GetPathToLogDir();
            string pathToLogFile = pathToLogDir + @"\log_" + DateTime.Now.ToString().Replace(".", "-").Replace(" ", "_").Replace(":", "-") + ".txt";
            bool exceptionWithLogsFlag = false;
            StringBuilder logsSTR = new StringBuilder();

            try
            {
                if (!Directory.Exists(pathToLogDir))
                {
                    exceptionWithLogsFlag = true;
                }

                logsSTR.AppendLine("----------------------------------------------------- TO DOCX SCRIPT -----------------------------------------------------");
                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - DomainName: " + Environment.UserDomainName + " ; UserName: " + Environment.UserName + " ; MachineName: " + Environment.MachineName);


                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("script:> Введите путь к рабочей директории");
                Console.Write("user:> ");
                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Введите путь к рабочей директории");
                logsSTR.Append("[ " + DateTime.Now.ToString() + " ] - user:> ");

                string pathToDir = Console.ReadLine();
                logsSTR.AppendLine(pathToDir);

                DirectoryInfo dirInfo = new DirectoryInfo(pathToDir);
                FileInfo[] wordFiles = null;

                if (dirInfo.Exists)
                {
                    bool check2 = true;

                    while (check2)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("script:> Укажите уровень просмотра (1 - без учёта поддиректорий, 2 - с учётом поддиректорий)");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Укажите уровень просмотра (1 - без учёта поддиректорий, 2 - с учётом поддиректорий)");
                        Console.Write("user:> ");
                        logsSTR.Append("[ " + DateTime.Now.ToString() + " ] - user:> ");
                        string f = Console.ReadLine();
                        logsSTR.AppendLine(f);
                        try
                        {
                            int respInt = Convert.ToInt32(f);
                            if (respInt == 1)
                            {
                                wordFiles = dirInfo.GetFiles("*.doc");
                                FileInfo[] rtfFiles = dirInfo.GetFiles("*.rtf");
                                wordFiles = wordFiles.Concat(rtfFiles).ToArray();

                                check2 = false;
                            }
                            else if (respInt == 2)
                            {
                                wordFiles = dirInfo.GetFiles("*.doc", SearchOption.AllDirectories);
                                FileInfo[] rtfFiles = dirInfo.GetFiles("*.rtf", SearchOption.AllDirectories);
                                wordFiles = wordFiles.Concat(rtfFiles).ToArray();

                                check2 = false;
                            }
                            else
                            {
                                throw new Exception();
                            }
                        }
                        catch
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("script:> Ошибка ввода");
                            logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Ошибка ввода");

                        }
                    }

                    bool checkACL = true;
                    while (checkACL)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("script:> Необходимо ли клонировать ACL (0 - нет, 1 - да)");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Необходимо ли клонировать ACL (0 - нет, 1 - да)");
                        Console.Write("user:> ");
                        logsSTR.Append("[ " + DateTime.Now.ToString() + " ] - user:>");

                        string resp = Console.ReadLine();
                        logsSTR.AppendLine(resp);

                        try
                        {
                            int respInt = Convert.ToInt32(resp);
                            if (respInt == 1)
                            {
                                copyACLFlag = true;
                                checkACL = false;
                            }
                            else if (respInt == 0)
                            {
                                copyACLFlag = false;
                                checkACL = false;
                            }
                            else
                            {
                                throw new Exception();
                            }
                        }
                        catch
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("script:> Ошибка ввода");
                            logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Ошибка ввода");
                        }
                    }

                    #region DialogToDeleteFile
                    bool check3 = true;
                    while (check3)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("script:> Необходимо ли удалить файлы DOC|RTF ? (0 - нет, 1 - да)");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Необходимо ли удалить файлы DOC|RTF ? (0 - нет, 1 - да)");
                        Console.WriteLine("user:> 1");
                        logsSTR.AppendLine("user:> 1");
                        string f = "1";

                        try
                        {
                            int respInt = Convert.ToInt32(f);
                            if (respInt == 0)
                            {
                                deleteFlag = false;
                                check3 = false;
                            }
                            else if (respInt == 1)
                            {
                                deleteFlag = true;
                                check3 = false;
                            }
                            else
                            {
                                throw new Exception();
                            }
                        }
                        catch
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("script:> Ошибка ввода");
                            logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Ошибка ввода");
                        }
                    }
                    #endregion

                    Console.WriteLine("script:> В рабочей директории обнаружено [" + wordFiles.Length + "] файла(ов) DOC|RTF:");
                    logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> В рабочей директории обнаружено [" + wordFiles.Length + "] файла(ов) DOC|RTF:");

                    int i = 1;
                    foreach (FileInfo wordFile in wordFiles)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                        Console.WriteLine("    [" + Convert.ToString(i) + "] - " + wordFile.FullName);
                        logsSTR.AppendLine("    [" + Convert.ToString(i) + "] - " + wordFile.FullName);
                        i++;
                    }


                    bool startFlag = false;
                    bool check = true;
                    while (check)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("script:> Запустить конвертацию в docx ? (0 - нет, 1 - да)");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Запустить конвертацию в docx ? (0 - нет, 1 - да)");
                        Console.Write("user:> ");
                        logsSTR.Append("[ " + DateTime.Now.ToString() + " ] - user:> ");
                        string resp = Console.ReadLine();
                        logsSTR.AppendLine(resp);
                        try
                        {
                            int respInt = Convert.ToInt32(resp);
                            if (respInt == 1)
                            {
                                startFlag = true;
                                check = false;
                            }
                            else if (respInt == 0)
                            {
                                Console.WriteLine("script:> Завершено!");
                                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Завершено!");
                                check = false;
                            }
                            else
                            {
                                throw new Exception();
                            }
                        }
                        catch
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("script:> Ошибка ввода");
                            logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Ошибка ввода");
                        }

                    }
                    if (startFlag)
                    {
                        Console.WriteLine("script:> Старт...");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Старт...");
                        word = new Microsoft.Office.Interop.Word.Application();
                        word.Visible = false;

                        object oMissing = System.Reflection.Missing.Value;

                        foreach (FileInfo wordFile in wordFiles)
                        {
                            try
                            {
                                Console.ForegroundColor = ConsoleColor.DarkGreen;
                                Console.Write("logs:> " + wordFile.FullName);
                                logsSTR.Append("[ " + DateTime.Now.ToString() + " ] - logs:> " + wordFile.FullName);

                                if (!wordFile.FullName.Contains("~$"))
                                {
                                    if (wordFile.Exists)
                                    {
                                        if (!fileIsOpen(wordFile.FullName))
                                        {
                                            Object filenamePath = (Object)wordFile.FullName;

                                            object confirmConversions = false;
                                            object oEncodDial = true;

                                            Microsoft.Office.Interop.Word.Document doc = word.Documents.OpenNoRepairDialog(
                                                ref filenamePath,
                                                ref confirmConversions,

                                                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                                                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                                                ref oMissing, ref oMissing, ref oEncodDial, oMissing);

                                            string newFileName = wordFile.FullName.Replace(".doc", ".docx");
                                            newFileName = newFileName.Replace(".rtf", ".docx");

                                            bool check4 = true;
                                            bool renamedFlag = false;
                                            int q = 0;
                                            while (check4)
                                            {
                                                FileInfo fileCheckExist = new FileInfo(newFileName);
                                                if (fileCheckExist.Exists)
                                                {
                                                    int index = newFileName.Length - 5;
                                                    newFileName = newFileName.Insert(index, "" + Convert.ToString(q) + "");
                                                    renamedFlag = true;
                                                    q++;
                                                }
                                                else
                                                {
                                                    check4 = false;
                                                }
                                            }


                                            doc.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                                            CompatibilityMode: WdCompatibilityMode.wdWord2010);
                                            ((Microsoft.Office.Interop.Word._Document)doc).Close();

                                            bool copyAclFlagException = false;
                                            string copyAclStrException = "";

                                            try
                                            {
                                                if (copyACLFlag)
                                                {
                                                    string script1 = @"$old = Get-Acl -path """ + wordFile.FullName + @"""";

                                                    string script2 = @"Set-Acl -path """ + newFileName + @""" -AclObject $old";
                                                    powerShell.AddScript(script1);
                                                    powerShell.AddScript(script2);
                                                    powerShell.Invoke();
                                                    powerShell.Commands.Clear();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                copyAclFlagException = true;
                                                copyAclStrException = ex.Message;
                                            }

                                            if (deleteFlag)
                                            {
                                                File.Delete(wordFile.FullName);
                                            }

                                            if (!renamedFlag)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Green;
                                                Console.Write("  [OK]");
                                                logsSTR.Append("  [OK]");
                                            }
                                            else
                                            {
                                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                                Console.Write("  [Renamed]");
                                                logsSTR.Append("  [Renamed]");
                                            }

                                            if (copyAclFlagException)
                                            {
                                                Console.ForegroundColor = ConsoleColor.DarkRed;
                                                Console.Write("  [ACL error]");
                                                logsSTR.Append("  [ACL error]");

                                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                                Console.WriteLine("script:> Не удалось клонировать ACL для данного файла { " + copyAclStrException + " }");

                                                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Не удалось клонировать ACL для данного файла { " + copyAclStrException + " }");
                                            }
                                            else
                                            {
                                                Console.ForegroundColor = ConsoleColor.Blue;
                                                Console.WriteLine("  [ACL cloned]");
                                                logsSTR.AppendLine("  [ACL cloned]");
                                            }

                                        }
                                        else
                                        {
                                            Console.ForegroundColor = ConsoleColor.DarkRed;
                                            Console.WriteLine("  [Blocked]");
                                            logsSTR.AppendLine("  [Blocked]");
                                        }
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                                        Console.WriteLine("  [NotFound]");
                                        logsSTR.AppendLine("  [NotFound]");
                                    }
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                                    Console.WriteLine("  [Warning]");
                                    logsSTR.AppendLine("  [Warning]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("logs:> " + ex.Message);
                                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - logs:> " + ex.Message);
                            }
                        }


                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("script:> Выполнено!");
                        logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - script:> Выполнено!");
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.ReadKey();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(ex.Message);
                logsSTR.AppendLine("[ " + DateTime.Now.ToString() + " ] - Критическая ошибка: { " + ex.Message + " }");
                Console.ReadKey();
            }
            finally
            {
                if (word != null)
                {
                    word.Quit();
                }

                if (!exceptionWithLogsFlag)
                {
                    using (FileStream stream = File.OpenWrite(pathToLogFile))
                    {
                        if (stream.CanWrite)
                        {
                            byte[] logs = Encoding.UTF8.GetBytes(logsSTR.ToString());
                            stream.Write(logs, 0, logs.Length);
                        }
                    }
                }
            }

        }

        static string GetPathToLogDir()
        {
            string path = ConfigurationManager.AppSettings.Get("PathToLogDir");
            if (string.IsNullOrEmpty(path))
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("script:> Введите путь для хранения файлов логов");
                Console.Write("user:> ");

                string userResp = Console.ReadLine();

                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine("script:> Для изменения пути хранения файлов логов перейдите в расположение \n\t скрипта и корректируйте раздел AppSettings файла ConvertDocToDocxScript.exe.config");

                if (Directory.Exists(userResp))
                {
                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    config.AppSettings.Settings["PathToLogDir"].Value = userResp;
                    try
                    {
                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("appSettings");
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("script:> Не удалось сохранить путь в файле конфигурации. Запустите скрипт \n\t с правами администратора или самостоятельно редактируйте раздел AppSettings \n\t файла ConvertDocToDocxScript.exe.config");
                    }
                    return userResp;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("script:> Не удалось получить доступ к указанной директории");
                    return "";
                }
            }
            else
            {
                return path;
            }
        }

        public static bool fileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);

                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }

                return false;
            }
            catch (System.IO.IOException ex)
            {
                return true;
            }
        }
    }
}
