using Microsoft.Office.Interop.Word;


OpenFileDialog file2 = new OpenFileDialog(); // sadece word için kullanılacak

private void DosyaSec_Click(object sender, EventArgs e) //dosya seç butonu
        {

            file2.Title = "Kelime Aranacak Dosyayı Seçiniz..";
            file2.Filter = @"Word Dosyaları (.docx ,.doc)|*.docx;*.doc";
            file2.FilterIndex = 1;
            file2.Multiselect = true;

            if (file2.ShowDialog() == DialogResult.OK)
            {
                // dosya seçildi ise


                string[] worddosyaIsimler = file2.SafeFileNames; //Çoklu seçimdeki dosyaların ismi
                string[] worddosyalar = file2.FileNames;
                string yaziolarak ="";
                for (int i = 0; i < worddosyaIsimler.Length; i++)
                {
                    yaziolarak = yaziolarak + worddosyaIsimler[i] + "\n";
                }

                MessageBox.Show("Seçilen Dosyalar : \n------------------------\n" + yaziolarak,
    "Seçim Özet");

                button5.BackColor = Color.Green;
                button5.ForeColor = Color.White;

            }

        }




private void PdfCevir_Click(object sender, EventArgs e) //Çevir

string[] worddosyalar = file2.FileNames;
            string[] wordisimler = file2.SafeFileNames;

            try
            {
                //Dosya Varmı Yokmu ? - Yoksa Oluştur..
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!Directory.Exists(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyalari\")))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path + @"\Yeni_Pdf_Dosyaları\"));

                }


                //PDF çevirme kodlarımız.
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();


                int i = 0;
                if (worddosyalar.Length > 0) //Dosya Seçilmiş mi ? 
                {


                    foreach (var dosya in worddosyalar)
                    {
                        wordDocument = appWord.Documents.Open(dosya);
                        wordDocument.ExportAsFixedFormat(path + @"\Yeni_Pdf_Dosyaları\" + wordisimler[i] + ".pdf",
                            WdExportFormat.wdExportFormatPDF);
                        i++;

                    }
                    MessageBox.Show("Dönüştürme Başarılı bir şekilde yapıldı\n Konum:\n\n" + path + @"\Yeni_Pdf_Dosyaları", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);


                }

                else
                {
                    MessageBox.Show("Lütfen ilk önce dosyaları seçin !", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            catch (Exception)
            {

                MessageBox.Show("Hata Oluştu \nDosya bozuk olabilir.\nDosya kullanılıyor olabilir.\nSeçtiğiniz dosyanın PDF'i zaten aynı konumda olabilir. \nTekrar Deneyin...\n\nCTRL Shift ve Esc Tuşlarına aynı anda basın\nTüm Word Dosyalarını Kapatın!",
                     "Uyarı",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Exclamation,
                     MessageBoxDefaultButton.Button1);
            }


        }
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
