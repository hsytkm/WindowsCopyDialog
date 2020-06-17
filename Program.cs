using Microsoft.VisualBasic.FileIO;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace WindowsCopyDialog
{
    // プロジェクトの "出力の種類" を "コンソールアプリケーション" から "Windowsアプリケーション" にすることで
    // コマンドプロンプト を表示されないようにした。
    class Program
    {
        static void Main(string[] args)
        {
            // 引数の処理
            if (args.Length < 2) throw new ArgumentOutOfRangeException("WindowsCopyDialog.exe [SourceFilePath] [DestinationFilePath]");
            var sourceFilePath = args[0];
            var destFilePath = args[1];

            if (sourceFilePath.ContainWildcard())
            {
                CopyWildcardFiles(sourceFilePath, destFilePath);
            }
            else
            {
                CopyFile(sourceFilePath, destFilePath);
            }
        }

        /// <summary>
        /// ワイルドカード含むファイルパスをコピー
        /// </summary>
        /// <param name="sourceFilePath">ワイルドカード含むファイルパス</param>
        /// <param name="destPath">コピーファイルPATH or コピー先ディレクトリPATH</param>
        /// <returns></returns>
        private static void CopyWildcardFiles(string sourceFilePath, string destPath)
        {
            if (!sourceFilePath.ContainWildcard()) throw new ArgumentException(sourceFilePath);

            var sourceFilenameRegex = Path.GetFileName(sourceFilePath).Replace(".", @"\.").Replace("*", @".+");

            var sourceDirPath = Path.GetDirectoryName(sourceFilePath);
            foreach (var filepath in Directory.EnumerateFiles(sourceDirPath))
            {
                if (Regex.IsMatch(Path.GetFileName(filepath), sourceFilenameRegex))
                    CopyFile(filepath, destPath);
            }
        }

        /// <summary>
        /// 
        /// https://docs.microsoft.com/ja-jp/dotnet/csharp/programming-guide/file-system/how-to-provide-a-progress-dialog-box-for-file-operations
        /// </summary>
        /// <param name="sourceFilePath"></param>
        /// <param name="destPath">コピーファイルPATH or コピー先ディレクトリPATH</param>
        /// <returns></returns>
        private static void CopyFile(string sourceFilePath, string destPath)
        {
            if (!File.Exists(sourceFilePath)) throw new FileNotFoundException(sourceFilePath);

            var destFilePath = File.Exists(destPath) ? destPath
                : Directory.Exists(destPath) ? Path.Combine(destPath, Path.GetFileName(sourceFilePath))
                : throw new DirectoryNotFoundException(destPath);

            if (File.Exists(destFilePath))
                FileSystem.DeleteFile(destFilePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);

            //Console.WriteLine($"copy {sourceFilePath} {destFilePath}");
            FileSystem.CopyFile(sourceFilePath, destFilePath, UIOption.AllDialogs);
        }

    }

    static class StringExtension
    {
        public static bool ContainWildcard(this string path) => path.Contains('*');
    }

}
