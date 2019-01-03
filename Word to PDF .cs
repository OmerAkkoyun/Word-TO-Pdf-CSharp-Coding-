string[] worddosyalar = file2.FileNames;
            string[] wordisimler = file2.SafeFileNames;

            try
            {
                //Dosya Varm� Yokmu ? - Yoksa Olu�tur..
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!Directory.Exists(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyalari\")))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyalar�\"));

                }


                //PDF �evirme kodlar�m�z.
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();


                int i = 0;
                if (worddosyalar.Length > 0) //Dosya Se�ilmi� mi ? 
                {


                    foreach (var dosya in worddosyalar)
                    {
                        wordDocument = appWord.Documents.Open(dosya);
                        wordDocument.ExportAsFixedFormat(path + @"\Yeni_Pdf_Dosyalar�\" + wordisimler[i] + ".pdf",
                            WdExportFormat.wdExportFormatPDF);
                        i++;

                        backgroundWorker2.ReportProgress(i);//progres bar'a rapor..
                    }
                    MessageBox.Show("D�n��t�rme Ba�ar�l� bir �ekilde yap�ld�\n Konum:\n\n" + path + @"\Yeni_Pdf_Dosyalar�", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);


                }

                else
                {
                    MessageBox.Show("L�tfen ilk �nce dosyalar� se�in !", "Uyar�", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            catch (Exception)
            {

                MessageBox.Show("Hata Olu�tu \nDosya bozuk olabilir.\nDosya kullan�l�yor olabilir.\nSe�ti�iniz dosyan�n PDF'i zaten ayn� konumda olabilir. \nTekrar Deneyin...\n\nCTRL Shift ve Esc Tu�lar�na ayn� anda bas�n\nT�m Word Dosyalar�n� Kapat�n!",
                     "Uyar�",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Exclamation,
                     MessageBoxDefaultButton.Button1);
            }


        }
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }