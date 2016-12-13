using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ris.Shuriken;

namespace WordMediaExtracter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            foreach (var s in args)
            {
                ExtractMediaFromDoc(s);
            }
        }

        private static void ExtractMediaFromDoc(string doc)
        {
            FolderSelectDialog dialog = new FolderSelectDialog
            {
                InitialDirectory = Path.GetDirectoryName(doc),
                Title = $@"Extract media from {Path.GetFileName(doc)} to:"
            };
            if (!dialog.Show()) return;

            string dir = dialog.FileName;
            Debug.WriteLine("--> {0}", (object) dir);

            try
            {
                using (var fs = File.OpenRead(doc))
                using (var archive = new ZipArchive(fs))
                {
                    var mediaEntries = archive.Entries.Where(entry => entry.IsInDir(@"word", @"media"));
                    bool any = false;
                    foreach (var zipArchiveEntry in mediaEntries)
                    {
                        Debug.WriteLine(zipArchiveEntry);
                        any = true;
                        zipArchiveEntry.ExtractToFile(Path.Combine(dir, zipArchiveEntry.Name));
                    }
                    if (any)
                    {
                        if (MessageBox.Show(@"All media has been extracted. Open folder?", @"Done!",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            Process.Start(dir);
                        }
                    }
                    else
                    {
                        MessageBox.Show(@"No media has been found.", @"Warning!", MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Could not extract media from file!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }

    internal static class Ext
    {
        internal static bool IsInDir(this ZipArchiveEntry entry, params string[] folders)
        {
            return entry.FullName.StartsWith(JoinAndAppend("/", folders)) ||
                   entry.FullName.StartsWith(JoinAndAppend("\\", folders));
        }

        internal static string JoinAndAppend(string sep, params string[] arr)
        {
            return string.Join(sep, arr) + sep;
        }
    }
}