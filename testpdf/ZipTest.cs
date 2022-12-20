using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using SharpCompress.Writers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace testpdf
{
    public class ZipTest
    {
        /// <summary>
        /// 将多个文件压缩到一个文件流中，可保存为zip文件，方便于web方式下载
        /// </summary>
        /// <param name="files">多个文件路径，文件或文件夹，第二个参数为重命名可包含路径</param>
        /// <param name="rootdir"></param>
        /// <returns>文件流</returns>
        public MemoryStream ZipStream(Dictionary<string, string> files, string rootdir = "")
        {
            using var archive = CreateZipArchive(files, rootdir);
            var ms = new MemoryStream();
            archive.SaveTo(ms, new WriterOptions(CompressionType.Deflate)
            {
                LeaveStreamOpen = true,
                ArchiveEncoding = new ArchiveEncoding()
                {
                    Default = Encoding.UTF8
                }
            });
            return ms;
        }

        /// <summary>
        /// 压缩多个文件
        /// </summary>
        /// <param name="files">多个文件路径，文件或文件夹</param>
        /// <param name="zipFile">压缩到...（包含文件名.zip）</param>
        /// <param name="rootdir">压缩包内部根文件夹</param>
        public void Zip(Dictionary<string, string> files, string zipFile, string rootdir = "")
        {
            using var archive = CreateZipArchive(files, rootdir);
            archive.SaveTo(zipFile, new WriterOptions(CompressionType.Deflate)
            {
                LeaveStreamOpen = true,
                ArchiveEncoding = new ArchiveEncoding()
                {
                    Default = Encoding.UTF8
                }
            });
        }

        /// <summary>
        /// 创建zip包
        /// </summary>
        /// <param name="files">文件路径</param>
        /// <param name="rootdir"></param>
        /// <returns></returns>
        private ZipArchive CreateZipArchive(Dictionary<string, string> files, string rootdir)
        {
            var archive = ZipArchive.Create();
            var dic = GetFileEntryMaps(files);
            //var remoteUrls = files.Distinct().Where(s => s.StartsWith("http")).Select(s =>
            //{
            //    try
            //    {
            //        return new Uri(s);
            //    }
            //    catch (UriFormatException)
            //    {
            //        return null;
            //    }
            //}).Where(u => u != null).ToList();
            foreach (var pair in dic)
            {
                archive.AddEntry(Path.Combine(rootdir, pair.Value), pair.Key);  // 给压缩文件重命名可以在这里修改，rootdir：文件夹路径 pair.Value：文件名（最好带后缀）
            }

            // 下面的是获取网络上的文件，暂时不需要
            //HttpClient _httpClient = new HttpClient();
            //if (remoteUrls.Any())
            //{
            //    var streams = new ConcurrentDictionary<string, Stream>();
            //    Parallel.ForEach(remoteUrls, url =>
            //    {
            //        _httpClient.GetAsync(url).ContinueWith(async t =>
            //        {
            //            if (t.IsCompleted)
            //            {
            //                var res = await t;
            //                if (res.IsSuccessStatusCode)
            //                {
            //                    Stream stream = await res.Content.ReadAsStreamAsync();
            //                    streams[Path.Combine(rootdir, Path.GetFileName(HttpUtility.UrlDecode(url.AbsolutePath)))] = stream;
            //                }
            //            }
            //        }).Wait();
            //    });
            //    foreach (var pair in streams)
            //    {
            //        archive.AddEntry(pair.Key, pair.Value);
            //    }
            //}

            return archive;
        }

        /// <summary>
        /// 获取文件路径和zip-entry的映射
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetFileEntryMaps(Dictionary<string, string> files)
        {
            var fileList = new Dictionary<string, string>();

            // 递归查找文件夹下的文件
            void GetFilesRecurs(string path)
            {
                //遍历目标文件夹的所有文件
                Directory.GetFiles(path).ToList().ForEach(filePath =>
                {
                    if (!fileList.ContainsKey(filePath)) // 避免下载同一个文件
                        fileList.Add(filePath, Path.GetFileName(filePath));
                });

                //遍历目标文件夹的所有文件夹
                foreach (var directory in Directory.GetDirectories(path))
                {
                    GetFilesRecurs(directory);
                }
            }

            foreach (var item in files)
            {
                if (Directory.Exists(item.Key))
                    GetFilesRecurs(item.Key); // 如果传文件夹，先不让他重命名
                else
                {
                    if (!fileList.ContainsKey(item.Key) && File.Exists(item.Key))
                        fileList.Add(item.Key, string.IsNullOrEmpty(item.Value) ? Path.GetFileName(item.Key) : item.Value);
                }
            }

            List<string> exportRename = new List<string>();
            // 导出名字重复处理
            foreach (var fileInfo in fileList)
                exportRename.Add(fileInfo.Value);

            if (exportRename.Any())
            {
                exportRename = exportRename.GroupBy(g => g).Where(q => q.Count() > 1).Select(q => q.Key).ToList();

                foreach (var fileRename in exportRename)
                {
                    var index = 0;
                    var dicRename = new Dictionary<string, string>();
                    var fileNameExitEtx = Path.GetFileNameWithoutExtension(fileRename);
                    var ext = Path.GetExtension(fileRename);
                    foreach (KeyValuePair<string, string> file in fileList)
                    {
                        if (file.Value == fileRename)
                        {
                            if (index > 0)
                            {
                                dicRename.Add(file.Key, $"{fileNameExitEtx}({index}){ext}");
                            }
                            index++;
                        }
                    }

                    foreach (var dic in dicRename)
                    {
                        fileList[dic.Key] = dic.Value;
                    }
                }
            }

            return fileList;
        }
    }
}
