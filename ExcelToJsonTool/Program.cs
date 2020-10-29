using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJsonTool
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                #region 将该文件夹下所有的txt 或者 csv 都转换为  json
                Console.WriteLine("请输入包含*.csv表格 的文件夹");
                string path = Console.ReadLine();

                string[] paths = GetAllFiles(path);

                if (paths != null)
                {
                    int length = paths.Length;
                    for (int i = 0; i < length; i++)
                    {
                        string currentPath = paths[i];
                        if (!currentPath.EndsWith(".txt") && !currentPath.EndsWith(".csv"))
                        {
                            continue;
                        }

                        string[] excel = File.ReadAllLines(currentPath, Encoding.Default);

                        List<string[]> table = ExcelToTable(excel);

                        string json = TableToJson(table);

                        string josnPath = currentPath + ".json";

                        File.WriteAllText(josnPath, json, Encoding.UTF8);

                        Console.WriteLine("已经将\r\n{0}\r\n保存到路径\r\n{1}\r\n", paths[i], josnPath);
                        //Console.WriteLine(json);
                    }
                }
                #endregion
            }
        }

        private static string[] GetAllDirectories(string path)
        {
            List<string> all = new List<string>();
            if (Directory.Exists(path))
            {
                string[] dirs = Directory.GetDirectories(path);
                int length = dirs.Length;
                for (int i = 0; i < length; i++)
                {
                    all.Add(dirs[i]);
                    all.AddRange(GetAllDirectories(dirs[i]));
                }
            }
            return all.ToArray();
        }

        private static string[] GetAllFiles(string path)
        {
            List<string> all = new List<string>();
            if (Directory.Exists(path))
            {
                Console.WriteLine("检索目录" + path);
                //当前文件夹下的所有文件
                all.AddRange(Directory.GetFiles(path));

                //当前文件夹下的所有文件夹
                string[] dirs = Directory.GetDirectories(path);
                int length = dirs.Length;
                for (int i = 0; i < length; i++)
                {
                    //遍历所有文件夹
                    all.AddRange(GetAllFiles(dirs[i]));

                }
            }
            return all.ToArray();
        }

        private static List<string[]> ExcelToTable(string text)
        {
            string[] excel = text.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            return ExcelToTable(excel);
        }

        private static List<string[]> ExcelToTable(string[] excel)
        {
            List<string[]> Table = new List<string[]>();

            int length = excel.Length;
            for (int i = 0; i < length; i++)
            {
                string[] columns = excel[i].Split("\t,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                Table.Add(columns);
            }
            return Table;
        }

        private static string TableToJson(List<string[]> table)
        {
            if (table == null || table.Count == 0)
            {
                return string.Empty;
            }
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.AppendLine("[");
            int count = table.Count;
            string[] head = table[0];
            for (int i = 1; i < count; i++)
            {
                string[] line = table[i];
                int length = line.Length;
                stringBuilder.AppendLine("{");
                for (int j = 0; j < length; j++)
                {
                    string key = head[j];
                    string value = line[j];
                    int intValue;
                    float flaotValue;
                    if(!int.TryParse(value,out intValue)&&!float.TryParse(value,out flaotValue))
                    {
                        value = "\"" + value + "\"";
                    }

                    if (j < head.Length)
                    {
                        stringBuilder.AppendFormat("\"{0}\":{1}", key, value);
                    }
                    else
                    {
                        stringBuilder.AppendFormat("\"{0}\":{1}", "未知", value);
                    }
                    if (j < length - 1)
                    {
                        stringBuilder.AppendLine(",");
                    }
                }
                
                stringBuilder.AppendLine("\r\n}");
                if (i < count - 1)
                {
                    stringBuilder.AppendLine(",");
                }
            }

            stringBuilder.AppendLine("]");

            return stringBuilder.ToString();
        }
        #region 没有使用的方法


        public static void ReadFile(string path, out byte[] bytes)
        {
            Console.WriteLine("开始读取");
            FileStream file = new FileStream(path, FileMode.OpenOrCreate);

            long length = file.Length;
            bytes = new byte[length];
            file.Read(bytes, 0, bytes.Length);
            file.Dispose();
            file.Close();
            Console.WriteLine(" 读取完毕");
        }

        public static void WriteFile(string path, byte[] bytes)
        {
            Console.WriteLine("开始写入");

            FileStream write = new FileStream(path, FileMode.OpenOrCreate);

            write.Write(bytes, 0, bytes.Length);
            write.Dispose();
            write.Close();
            Console.WriteLine(" 读取写入完毕");
        }

        public static void CopyFile(string readPath, string writePath)
        {
            byte[] bytes;

            ReadFile(readPath, out bytes);

            WriteFile(writePath, bytes);
        }

        public static string ReadText(string path, Encoding encoding)
        {
            byte[] bytes;
            ReadFile(path, out bytes);

            return encoding.GetString(bytes);
        }

        public static void WriteText(string path, string text, Encoding encoding)
        {
            byte[] bytes = encoding.GetBytes(text);
            WriteFile(path, bytes);
            Console.WriteLine("文本文件写入完毕");
        }
        #endregion
    }
}
