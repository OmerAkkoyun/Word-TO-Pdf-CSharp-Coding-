string[] worddosyalar = file2.FileNames;
            string[] wordisimler = file2.SafeFileNames;

            try
            {
                //Dosya Varmý Yokmu ? - Yoksa Oluþtur..
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!Directory.Exists(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyalari\")))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyalarý\"));

                }


                //PDF çevirme kodlarýmýz.
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();


                int i = 0;
                if (worddosyalar.Length > 0) //Dosya Seçilmiþ mi ? 
                {


                    foreach (var dosya in worddosyalar)
                    {
                        wordDocument = appWord.Documents.Open(dosya);
                        wordDocument.ExportAsFixedFormat(path + @"\Yeni_Pdf_Dosyalarý\" + wordisimler[i] + ".pdf",
                            WdExportFormat.wdExportFormatPDF);
                        i++;

                        backgroundWorker2.ReportProgress(i);//progres bar'a rapor..
                    }
                    MessageBox.Show("Dönüþtürme Baþarýlý bir þekilde yapýldý\n Konum:\n\n" + path + @"\Yeni_Pdf_Dosyalarý", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);


                }

                else
                {
                    MessageBox.Show("Lütfen ilk önce dosyalarý seçin !", "Uyarý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            catch (Exception)
            {

                MessageBox.Show("Hata Oluþtu \nDosya bozuk olabilir.\nDosya kullanýlýyor olabilir.\nSeçtiðiniz dosyanýn PDF'i zaten ayný konumda olabilir. \nTekrar Deneyin...\n\nCTRL Shift ve Esc Tuþlarýna ayný anda basýn\nTüm Word Dosyalarýný Kapatýn!",
                     "Uyarý",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Exclamation,
                     MessageBoxDefaultButton.Button1);
            }


        }
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }