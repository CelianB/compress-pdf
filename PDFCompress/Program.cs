using Spire.Pdf;
using Spire.Pdf.Exporting;
using Spire.Pdf.Graphics;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace PDFCompress
{
    class Program
    {
        static void Main(string[] args)
        {
            var isGoodPath = false;
            while (!isGoodPath)
            {
                Console.Write("Folder containing pdf to reduce :");
                var path = Console.ReadLine();
                if (Directory.Exists(path))
                {
                    Directory.CreateDirectory(path + "/temp");
                    isGoodPath = true;
                    Console.WriteLine("File compression in progress ...");
                    int i = 0;
                    var files = new List<string> (Directory.GetFiles(path, "*.pdf"));

                    Parallel.ForEach(files, item =>
                    {
                        i++;
                        var name = Path.GetFileName(item);
                        var dir = Path.GetDirectoryName(item);

                        var dest = dir + @"\temp\" + name;
                        if (!File.Exists(dest))
                        {
                            Console.WriteLine($"Compression of {name} {i}/{files.Count} (copy in {dir}\temp)");
                            CompressPdf(item, dest);
                        }
                        else Console.WriteLine($"Existing file : {name}");
                    });
                }
                else
                    Console.WriteLine("Invalid path");
            }
        }

        static void CompressPdf(string source, string dest)
        {
            var doc = new PdfDocument(source);

                //Disable incremental update
                doc.FileInfo.IncrementalUpdate = false;
                doc.CompressionLevel = PdfCompressionLevel.Best;
                //doc.PrintSettings.Color = false;
                //Traverse all pages of PDF and diagonse whether the images are contained or not
                foreach (PdfPageBase page in doc.Pages)
                {
                    if (page != null)
                    {

                        Image[] images = page.ExtractImages();
                        //Traverse all images
                        if (images != null && images.Length > 0)
                        {
                            for (int j = 0; j < images.Length; j++)
                            {
                                Image image = images[j];
                                PdfBitmap bp = new PdfBitmap(image);
                                //Set bp.Quality values to compress images
                                bp.Quality = 10;
                                //Replace the images with newly compressed images
                                page.ReplaceImage(j, bp);
                            }
                        }

                        if (page.ImagesInfo != null)
                        {
                            foreach (PdfImageInfo info in page.ImagesInfo)
                            {
                                //Call the method TryCompressImage() to compress the images in PDF 
                                page.TryCompressImage(info.Index);
                            }
                        }
                    }
                }
                //Save to file
                doc.SaveToFile(dest);
                doc.Close();
                doc.Dispose();
            }
        
    }
}
