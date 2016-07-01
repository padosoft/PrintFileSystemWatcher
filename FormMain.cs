using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Management;
using System.Diagnostics;
using System.Drawing.Printing;
using FileSystemWatcher;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Widget;
using System.IO.Compression;

namespace FileChangeNotifier
{
    public partial class frmNotifier : Form
    {
        private StringBuilder m_Sb;
        private bool m_bDirty;
        private System.IO.FileSystemWatcher m_Watcher;
        private bool m_bIsWatching;

        private long cntList = 0;


        bool UseRawHelper = false;
		Font printFont = new Font("Arial", 10);
		System.IO.StreamReader streamToPrint = null;
		//string EtichettaTXTPath = Application.StartupPath+@"\etichetta_zebra.txt";
		//string PrinterName = @"zebra";
        System.Drawing.Printing.PrintDocument printDocument1 ;

        public frmNotifier()
        {
            InitializeComponent();
            m_Sb = new StringBuilder();
            m_bDirty = false;
            m_bIsWatching = false;

            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                listPrinters.Items.Add(printer);
                
                string dir = Application.StartupPath + @"\printers\" + printer.Replace(@"\", @"%");
                if(!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }

            if (!Directory.Exists(Application.StartupPath + @"\printers\zebra_tcp"))
            {
                Directory.CreateDirectory(Application.StartupPath + @"\printers\zebra_tcp");
            }
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(OnFormClosing);

            txtFile.Text = Application.StartupPath + @"\printers\";
            
            /*var printerQuery = new ManagementObjectSearcher("SELECT * from Win32_Printer");
            foreach (var printer in printerQuery.Get())
            {
                var name = printer.GetPropertyValue("Name");
                var status = printer.GetPropertyValue("Status");
                var isDefault = printer.GetPropertyValue("Default");
                var isNetworkPrinter = printer.GetPropertyValue("Network");

                listPrinters.Items.Add(name + " (Status: " + status + ", Default: " + isDefault + ", Network: " + isNetworkPrinter + "");
            }*/

        }

        

            private void OnFormClosing(object sender, FormClosingEventArgs e)
            {
                if (DialogResult.No == MessageBox.Show("Are you sure you want to close Application?", "Attention!", MessageBoxButtons.YesNo)) e.Cancel = true;
            }

        private void btnWatchFile_Click(object sender, EventArgs e)
        {
            if (m_bIsWatching)
            {
                m_bIsWatching = false;
                m_Watcher.EnableRaisingEvents = false;
                m_Watcher.Dispose();
                btnWatchFile.BackColor = Color.LightSkyBlue;
                btnWatchFile.Text = "Start Watching";
                
            }
            else
            {
                m_bIsWatching = true;
                btnWatchFile.BackColor = Color.Red;
                btnWatchFile.Text = "Stop Watching";

                m_Watcher = new System.IO.FileSystemWatcher();
                if (rdbDir.Checked)
                {
                    m_Watcher.Filter = "*.*";
                    m_Watcher.Path = txtFile.Text ;
                }
                else
                {
                    m_Watcher.Filter = txtFile.Text.Substring(txtFile.Text.LastIndexOf('\\') + 1);
                    m_Watcher.Path = txtFile.Text.Substring(0, txtFile.Text.Length - m_Watcher.Filter.Length);
                }

                if (chkSubFolder.Checked)
                {
                    m_Watcher.IncludeSubdirectories = true;
                }

                m_Watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                                     | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                m_Watcher.Changed += new FileSystemEventHandler(OnChanged);
                m_Watcher.Created += new FileSystemEventHandler(OnChanged);
                m_Watcher.Deleted += new FileSystemEventHandler(OnChanged);
                m_Watcher.Renamed += new RenamedEventHandler(OnRenamed);
                m_Watcher.EnableRaisingEvents = true;
            }
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (!m_bDirty)
            {
                m_Sb.Remove(0, m_Sb.Length);
                m_Sb.Append(e.FullPath);
                m_Sb.Append(" ");
                m_Sb.Append(e.ChangeType.ToString());
                m_Sb.Append("    ");
                m_Sb.Append(DateTime.Now.ToString());
                m_bDirty = true;
                count();
                if(e.ChangeType== WatcherChangeTypes.Created)
                    {
                        if (e.FullPath.Substring(e.FullPath.Length-3)=="pdf")
                            {
                                PrintFilePdf(e.FullPath, Path.GetDirectoryName(e.FullPath).Replace(@"%", @"\").Substring(m_Watcher.Path.Length));
                                File.Delete(e.FullPath);
                                return;
                            }
                        if (e.FullPath.Substring(e.FullPath.Length - 3) == "txt" && e.FullPath.Contains("zebra_tcp"))
                            {
                                //PrintFilePdf(e.FullPath, Path.GetDirectoryName(e.FullPath).Replace(@"%", @"\").Substring(m_Watcher.Path.Length));
                                string[] ip_port_stampante = Path.GetDirectoryName(e.FullPath).Replace(@"%", @"\").Substring(m_Watcher.Path.Length + 10).Split('\\');
                                int port = 0;
                                if (!int.TryParse(ip_port_stampante[1],out port))
                                    {
                                        writeLog("formattazione numero porta stampante tcp zebra errato");
                                    }
                                stampa(ip_port_stampante[0], port, e.FullPath, ip_port_stampante[2]);
                                File.Delete(e.FullPath);
                                return;
                            }
                        if (e.FullPath.Substring(e.FullPath.Length - 3) == "txt" )
                        {
                            //PrintFilePdf(e.FullPath, Path.GetDirectoryName(e.FullPath).Replace(@"%", @"\").Substring(m_Watcher.Path.Length));
                            stampa(null, 0, e.FullPath, Path.GetDirectoryName(e.FullPath).Replace(@"%", @"\").Substring(m_Watcher.Path.Length));
                            File.Delete(e.FullPath);
                            return;
                        }
                    }

            }

        }



        private void OnRenamed(object sender, RenamedEventArgs e)
        {
            if (!m_bDirty)
            {
                m_Sb.Remove(0, m_Sb.Length);
                m_Sb.Append(e.OldFullPath);
                m_Sb.Append(" ");
                m_Sb.Append(e.ChangeType.ToString());
                m_Sb.Append(" ");
                m_Sb.Append("to ");
                m_Sb.Append(e.Name);
                m_Sb.Append("    ");
                m_Sb.Append(DateTime.Now.ToString());
                m_bDirty = true;
                count();
                if (rdbFile.Checked)
                {
                    m_Watcher.Filter = e.Name;
                    m_Watcher.Path = e.FullPath.Substring(0, e.FullPath.Length - m_Watcher.Filter.Length);
                }
            }            
        }

        private void count()
            {
                cntList++;
                System.Threading.Thread.Sleep(500);
                writeLog(lstNotification.Items[lstNotification.Items.Count-1].ToString());
                if (cntList > 100)
                {
                        lstNotification.Invoke((MethodInvoker)(() => lstNotification.Items.Clear()));
                        cntList = 0;
                }
            }

        private void tmrEditNotify_Tick(object sender, EventArgs e)
        {
            if (m_bDirty)
            {
                lstNotification.BeginUpdate();
                lstNotification.Items.Add(m_Sb.ToString());
                lstNotification.EndUpdate();
                m_bDirty = false;
            }
        }

        private void btnBrowseFile_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", txtFile.Text);
            /*if (rdbDir.Checked)
            {
                DialogResult resDialog = dlgOpenDir.ShowDialog();
                if (resDialog.ToString() == "OK")
                {
                    txtFile.Text = dlgOpenDir.SelectedPath;
                }
            }
            else
            {
                DialogResult resDialog = dlgOpenFile.ShowDialog();
                if (resDialog.ToString() == "OK")
                {
                    txtFile.Text = dlgOpenFile.FileName;
                }
            }*/
        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            DialogResult resDialog = dlgSaveFile.ShowDialog();
            if (resDialog.ToString() == "OK")
            {
                FileInfo fi = new FileInfo(dlgSaveFile.FileName);
                StreamWriter sw = fi.CreateText();
                foreach (string sItem in lstNotification.Items)
                {
                    sw.WriteLine(sItem);
                }
                sw.Close();
            }
        }

        private void rdbFile_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbFile.Checked == true)
            {
                chkSubFolder.Enabled = false;
                chkSubFolder.Checked = false;
            }
        }

        private void rdbDir_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbDir.Checked == true)
            {
                chkSubFolder.Enabled = true;
            }
        }



        private void PrintFilePdf(string filename, string printername)
            {
                    if(File.Exists(filename))
                        {
                            PdfDocument doc = new PdfDocument();
                            doc.LoadFromFile(filename);
                            PrintDialog dialogPrint = new PrintDialog();
                            dialogPrint.AllowPrintToFile = true;
                            dialogPrint.AllowSomePages = true;
                            dialogPrint.PrinterSettings.MinimumPage = 1;
                            dialogPrint.PrinterSettings.MaximumPage = doc.Pages.Count;
                            dialogPrint.PrinterSettings.FromPage = 1;
                            dialogPrint.PrinterSettings.ToPage = doc.Pages.Count;


                            //Set the pagenumber which you choose as the start page to print
                            doc.PrintFromPage = dialogPrint.PrinterSettings.FromPage;
                            //Set the pagenumber which you choose as the final page to print
                            doc.PrintToPage = dialogPrint.PrinterSettings.ToPage;
                            //Set the name of the printer which is to print the PDF
                            doc.PrinterName = printername;

                            PrintDocument printDoc = doc.PrintDocument;
                            dialogPrint.Document = printDoc;
                            printDoc.Print();
                        }
                
            }

        private bool printerExist(string printername)
            {
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                    {
                        if (printer == printername) return true;
                    }
                return false;
            }

        private void writeLog(string logtext)
        {
            FileInfo fi = new FileInfo("log.txt");
            if (File.Exists("log.txt") && fi.Length > 1000000) 
                {
                Compress(fi);
                DateTime ora = DateTime.Now;
                File.Move("log.txt.gz", "log.txt." + ora.Year + "." + ora.Month + "." + ora.Day + "." + ora.Hour + "." + ora.Minute + "." + ora.Second + ".gz");
                File.Delete("log.txt");
                }
            StreamWriter sw;
            if (File.Exists("log.txt")) sw = new StreamWriter("log.txt",true);
            else sw = fi.CreateText();
            sw.WriteLine(logtext);
            sw.Close();
        }


        public void stampa(string ipAddress, int TCPIPPort, string filename, string PrinterName)
{
			try
			{
				if (ipAddress != null && !string.IsNullOrEmpty(ipAddress) && TCPIPPort > 0)
				{
                    PrintZebraFileToTCPIP(ipAddress, TCPIPPort, filename);
				}
				else if (!UseRawHelper)
				{
                    using (streamToPrint = new System.IO.StreamReader(filename))
					{
						try
						{
							printFont = new Font("Arial", 10);
							printDocument1 = new System.Drawing.Printing.PrintDocument();
							printDocument1.PrinterSettings.PrinterName = PrinterName;
							printDocument1.PrinterSettings.Copies = 1;
							printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocument1_PrintPage);
							printDocument1.Print();
						}
						finally
						{
							try
							{
								streamToPrint.Close();
							}
							catch { }
						}
					}
				}
				else
				{
                    if (!RawPrinterHelper.RawPrinterHelper.SendFileToPrinter(PrinterName, filename))
					{
						writeLog("ERRORE DURANTE L'INVIO DEL FILE ZEBRA ALLA STAMPANTE!");
						//return;
					}
				}
				System.Threading.Thread.Sleep(1000);
			}
			catch (Exception ex)
			{
				writeLog("ERRORE DURANTE LA STAMPA!" + Environment.NewLine + Environment.NewLine + "Messaggio di errore:" + Environment.NewLine + ex.Message);
				return;
			}
}





		private void PrintZebraFileToTCPIP(string ipAddress, int port, string EtichettaTXTPath)
		{
			try
			{
				writeLog("PrintZebraFileToTCPIP: ipAddress: " + ipAddress + " - port: " + port.ToString());
			}
			catch (Exception ex)
			{
				writeLog("ERROR in PrintZebraFileToTCPIP() durante il Log: \r\n\r\n" + ex.Message);
			}

			string ZPLString = "";
			using (streamToPrint = new System.IO.StreamReader(EtichettaTXTPath))
			{
				try
				{
					ZPLString = streamToPrint.ReadToEnd();
				}
				finally
				{
					try
					{
						streamToPrint.Close();
					}
					catch { }
				}
			}
			if (String.IsNullOrEmpty(ZPLString))
			{
				writeLog("ERRORE DURANTE LA LETTURA DEL FILE ETICHETTA: "+EtichettaTXTPath );
			}
			else
			{
				PrintZebraZPLToTCPIP(ipAddress, port, ZPLString);
			}
		}

		private void PrintZebraZPLToTCPIP(string ipAddress, int port, string ZPLString)
		{
			try
			{
				// Open connection
				System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();
				client.Connect(ipAddress, port);

				// Write ZPL String to connection
				System.IO.StreamWriter writer = new System.IO.StreamWriter(client.GetStream());
				writer.Write(ZPLString);
				writer.Flush();

				// Close Connection
				writer.Close();
				client.Close();
			}
			catch (Exception ex)
			{
				writeLog("ERRORE DURANTE LA STAMPA IN TCP/IP!" + Environment.NewLine + Environment.NewLine + "Messaggio di errore:" + Environment.NewLine + ex.Message);
			}		
		}





		void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ev)
		{
			Single linesPerPage = 0;
			Single yPos = 0;
			int count = 0;
			Single leftMargin = ev.MarginBounds.Left;
			Single topMargin = ev.MarginBounds.Top;
			string line = "";

			// Calculate the number of lines per page.
			linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);

			// Print each line of the file.
			while (count < linesPerPage)
			{
				line = streamToPrint.ReadLine();
				if (line == null)
				{
					break;
				}
				yPos = topMargin + count * printFont.GetHeight(ev.Graphics);
				ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, yPos, new StringFormat());
				count++;
			}

			// If more lines exist, print another page.
			if (line != null)
			{
				ev.HasMorePages = true;
			}
			else
			{
				ev.HasMorePages = false;
			}
		}

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void frmNotifier_Load(object sender, EventArgs e)
        {
            btnWatchFile.PerformClick();
        }

        public static void Compress(FileInfo fileToCompress)
        {

                using (FileStream originalFileStream = fileToCompress.OpenRead())
                {
                    if ((File.GetAttributes(fileToCompress.FullName) &
                       FileAttributes.Hidden) != FileAttributes.Hidden & fileToCompress.Extension != ".gz")
                    {
                        using (FileStream compressedFileStream = File.Create(fileToCompress.FullName + ".gz"))
                        {
                            using (GZipStream compressionStream = new GZipStream(compressedFileStream,
                               CompressionMode.Compress))
                            {
                                originalFileStream.CopyTo(compressionStream);

                            }
                        }
                        //DateTime ora = DateTime.Now;

                        FileInfo info = new FileInfo(fileToCompress.Name + ".gz");
                        Console.WriteLine("Compressed {0} from {1} to {2} bytes.",
                        fileToCompress.Name, fileToCompress.Length.ToString(), info.Length.ToString());
                        //File.Move(fileToCompress.Name, fileToCompress.Name.Substring(0, fileToCompress.Name.Length - 3) + "." + ora.Year + "." + ora.Month + "." + ora.Day + "." + ora.Hour + "." + ora.Minute + "." + ora.Second + ".gz");
                    }

                }
            
        }

        private void frmNotifier_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
                Hide();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void readmeZebraTCPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", "readme.txt");
        }

	}

    }

