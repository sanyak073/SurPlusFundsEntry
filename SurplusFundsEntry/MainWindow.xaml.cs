using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;
using Winforms = System.Windows.Forms;
using iTextSharp.text.pdf;
using System.Xml.Serialization;


namespace SurplusFundsEntry
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	/// 

	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			mtgDate.SelectedDate = DateTime.Now;
			judDateRec.SelectedDate = DateTime.Now;
			dateReviewed.SelectedDate = DateTime.Now;
			dateRecorded.SelectedDate = DateTime.Now;
			dateForclosed.SelectedDate = DateTime.Now;
			conclusionDate.SelectedDate = DateTime.Now;
			dtEstateDateReviewed.SelectedDate = DateTime.Now;

			comboBox1.SelectedIndex = 0;
		}

		public static List<Tuple<string, string>> CountyStateList = new List<Tuple<string, string>>();
		public static string FileToOpen = "";

		public static string firstPagesMtgPDF = "";
		public static string firstPagesTaxPDF = "";
		public static string mtgPagesPDF = "";
		public static string judgmentpagesPDF = "";
		public static string conclusionPagesPDF = "";
		public static string estatePagePdf = "";

		public static bool mtgExtended = false;
		public static bool judgmentExtended = false;
		public static bool conclusionExtended = false;



		public static string pdfSavePath;
		public static string xlSavePath;

		public static string xlTemplatePath;
		public static string documentDir;


		public static int mtgID;
		public static int judgementID;
		public static int claimantID;
		public static int conclusionID;

		public static int selectedID;

		public static string satisfied;
		public static string mtgSurplusAmount = "";

		public static string researcher;

		public static Boolean pdfCheck1 = false;
		public static Boolean pdfCheck2 = false;
		public static Boolean pdfCheck3 = false;
		public static Boolean pdfCheck4 = false;

		public static Boolean pdfCheck5 = false;
		public static Boolean pdfCheck6 = false;
		public static Boolean pdfMisc = false;

		//not used
		public static Boolean estatesCheck1 = false;
		public static Boolean estatesCheck2 = false;
		public static Boolean estatesCheck3 = false;
		public static Boolean estatesCheck4 = false;
		public static Boolean estatesCheck5 = false;
		//--

		public static Boolean hasEstatesForm = false;

		public static Boolean isMtgForm = false;
		public static Boolean isTaxForm = false;
		public static Boolean isDeedsForm = false;

		public static Boolean isMtgLoaded = false;
		public static Boolean isJudLoaded = false;
		public static Boolean isConcLoaded = false;
		public static Boolean isClaimantLoaded = false;

		public static List<mtgClass> mtgList = new List<mtgClass>();
		public static List<judgementClass> judList = new List<judgementClass>();
		public static List<claimantsClass> claimantList = new List<claimantsClass>();
		public static List<conclusionClass> conclusionList = new List<conclusionClass>();

		public static List<string> docs = new List<string>();
		public static List<string> anyDocsList = new List<string>();
		public static List<string> anyOtherDocsList = new List<string>();

		List<string> tempMisc = new List<string>();

		project1 p1 = new project1();

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			ScrollViewer.Visibility = System.Windows.Visibility.Hidden;
			p1 = new project1();
			ClearStatic();
			Retry:
			if (!File.Exists(Properties.Settings.Default.pathTemplate))
			{
				MessageBox.Show("Missing path to Excel TEMPLATE. Please select the file");
				OpenFileDialog opendialog = new OpenFileDialog();
				opendialog.Filter = "Excel files|*.xl*";
				opendialog.Title = "EXCEL TEMPLATE";
				if (opendialog.ShowDialog() == true)
				{
					Properties.Settings.Default.pathTemplate = opendialog.FileName;
					Properties.Settings.Default.Save();
				}
				else
				{
					MessageBox.Show("No files selected. File path is required!");
					goto Retry;
				}

			}

			if (!File.Exists(Properties.Settings.Default.pathDeeds))
			{
				MessageBox.Show("Missing path to Excel DEEDS. Please select the file");
				OpenFileDialog opendialog = new OpenFileDialog();
				opendialog.Filter = "Excel files|*.xl*";
				opendialog.Title = "EXCEL DEEDS";
				if (opendialog.ShowDialog() == true)
				{
					Properties.Settings.Default.pathDeeds = opendialog.FileName;
					Properties.Settings.Default.Save();
				}
				else
				{
					MessageBox.Show("No files selected. File path is required!");
					goto Retry;
				}
			}

			if (!Directory.Exists(Properties.Settings.Default.pathSaveFolder))
			{
				MessageBox.Show("Missing folder path for <EXCEL> saving. Please select the folder.");
				Winforms.FolderBrowserDialog fdiag = new Winforms.FolderBrowserDialog();
				fdiag.Description = "EXCEL SAVE FOLDER";

				if (fdiag.ShowDialog() == Winforms.DialogResult.OK)
				{
					Properties.Settings.Default.pathSaveFolder = fdiag.SelectedPath;
					Properties.Settings.Default.Save();
				}
				else
				{
					MessageBox.Show("Path to <EXCEL> save folder is not set. Path required!");
					goto Retry;
				}

			}

			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWb = xlApp.Workbooks.Open(Properties.Settings.Default.pathTemplate, false, false);
			Excel.Worksheet xlWs = xlWb.Worksheets["countylist"];

			for (int rowx = 2; rowx <= xlWs.UsedRange.Rows.Count; rowx++)
			{
				if (!tbState.Items.Contains(xlWs.Cells[rowx, 2].Value2))
				{
					tbState.Items.Add(xlWs.Cells[rowx, 2].Value2);
					cbEstatesState.Items.Add(xlWs.Cells[rowx, 2].Value2);
				}
				CountyStateList.Add(new Tuple<string, string>(xlWs.Cells[rowx, 1].value, xlWs.Cells[rowx, 2].Value2));
				if (!tbHeld.Items.Contains(xlWs.Cells[rowx, 3].Value2))
					tbHeld.Items.Add(xlWs.Cells[rowx, 3].Value2);
			}
			try
			{
				tbHeld.SelectedIndex = 0;
			}
			catch
			{

			}

			xlWb.Close(false);
			xlApp.Quit();
			mtgID = 101;
			mtgLabelCount.Content = "0 OF 13";
			judgementID = 201;
			judgementLabelCount.Content = "0 OF 12";

			claimantID = 1;
			claimantsLabelCount.Content = "0 OF 7";

			conclusionID = 301;

			tbResearcher.SelectedIndex = 0;
			tbEstatesResearcher.Text = tbResearcher.Text;

			tbHeld.SelectedIndex = 0;
			tbEstatesHow.Text = tbHeld.Text;

			tbName.Focus();

			if (isTaxForm)
			{
				lblmtgDateForclosed.Margin = lblmtgDateRec.Margin;
				lblmtgDateRec.Visibility = System.Windows.Visibility.Collapsed;

				dateForclosed.Margin = dateRecorded.Margin;
				dateRecorded.Visibility = System.Windows.Visibility.Collapsed;

				lblmtgBook.Visibility = System.Windows.Visibility.Collapsed;
				lblmtgPage.Visibility = System.Windows.Visibility.Collapsed;

				tbBook.Visibility = System.Windows.Visibility.Collapsed;
				tbPage.Visibility = System.Windows.Visibility.Collapsed;

				lblmtgAmountNote.Content = "AMOUNT OF FORECLOSED ON(JUDGEMENT AMOUNT)";
				lblmtgBeneficiary.Content = "FORECLOSING ENTITY (COUNTRY OR CITY NAME)";

				lblmtgAmountNote.Margin = new Thickness(lblmtgAmountNote.Margin.Left, lblmtgAmountNote.Margin.Top - 70, lblmtgAmountNote.Margin.Right, lblmtgAmountNote.Margin.Bottom);
				lblmtgBeneficiary.Margin = new Thickness(lblmtgBeneficiary.Margin.Left, lblmtgBeneficiary.Margin.Top - 70, lblmtgBeneficiary.Margin.Right, lblmtgBeneficiary.Margin.Bottom);

				tbNote.Margin = new Thickness(tbNote.Margin.Left, tbNote.Margin.Top - 70, tbNote.Margin.Right, tbNote.Margin.Bottom);
				tbBeneficiary.Margin = new Thickness(tbBeneficiary.Margin.Left, tbBeneficiary.Margin.Top - 70, tbBeneficiary.Margin.Right, tbBeneficiary.Margin.Bottom);

			}

			expander8.Visibility = System.Windows.Visibility.Collapsed;


		}

		void ClearStatic()
		{
			CountyStateList = new List<Tuple<string, string>>();

			firstPagesMtgPDF = "";
			firstPagesTaxPDF = "";
			mtgPagesPDF = "";
			judgmentpagesPDF = "";
			conclusionPagesPDF = "";
			estatePagePdf = "";

			mtgExtended = false;
			judgmentExtended = false;
			conclusionExtended = false;


			pdfSavePath = "";
			xlSavePath = "";

			xlTemplatePath = "";
			documentDir = "";

			mtgID = 101;
			judgementID = 201;
			claimantID = 1;
			conclusionID = 301;


			satisfied = "";

			mtgSurplusAmount = "";

			pdfCheck1 = false;
			pdfCheck2 = false;
			pdfCheck3 = false;
			pdfCheck4 = false;

			pdfCheck5 = false;
			pdfCheck6 = false;
			pdfMisc = false;



			hasEstatesForm = false;

			isMtgLoaded = false;
			isJudLoaded = false;
			isConcLoaded = false;
			isClaimantLoaded = false;

			mtgList = new List<mtgClass>();
			judList = new List<judgementClass>();
			claimantList = new List<claimantsClass>();
			conclusionList = new List<conclusionClass>();

		}

		private void button3_Click(object sender, RoutedEventArgs e)
		{
			expander1.IsExpanded = false;
			expander2.IsExpanded = true;
		}
		private void tbSurplus_TextChanged(object sender, TextChangedEventArgs e)
		{
			if (tbSurplus.Text.Length > 0)
			{
				double res;
				if (double.TryParse(tbSurplus.Text, out res) == false)
				{
					MessageBox.Show("The value is not in the correct format!");
				}
				tbEstatesFunds.Text = tbSurplus.Text;
			}
		}

		private void button4_Click(object sender, RoutedEventArgs e)
		{
			string pdf = mtgAttButton.Content.ToString().Replace("ATTACHMENT (PDF)", "");
			ComboBoxItem cbItem = (ComboBoxItem)comboBox1.SelectedItem;

			if (isMtgLoaded)
			{
				if (cbItem.Content.ToString() == "-")
				{
					mtgBookSatisfied.Text = "";
					mtgPageSatisfied.Text = "";
					DeleteConclusion(selectedID);

					AddConclusion(selectedID, mtgName.Text, mtgAmount.Text, mtgDate.SelectedDate.Value, pdf, false);
				}
				else
				{
					try
					{
						DeleteConclusion(selectedID);
					}
					catch
					{

					}
				}

				DeleteMtg(selectedID);
				AddMtgs(selectedID, mtgName.Text, mtgAmount.Text, mtgLoanNum.Text, mtgDate.SelectedDate.Value, mtgBook.Text, mtgPage.Text, mtgAddress.Text, satisfied, pdf, mtgBookSatisfied.Text, mtgPageSatisfied.Text, comboBox1.Text);
				isMtgLoaded = false;
			}

			else
			{
				while (mtgList.Find(x => x.ID == mtgID) != null)
				{
					mtgID++;
				}

				if (cbItem.Content.ToString() == "-")
				{
					mtgBookSatisfied.Text = "";
					mtgPageSatisfied.Text = "";

					AddConclusion(mtgID, mtgName.Text, mtgAmount.Text, mtgDate.SelectedDate.Value, pdf, false);

				}



				AddMtgs(mtgID, mtgName.Text, mtgAmount.Text, mtgLoanNum.Text, mtgDate.SelectedDate.Value, mtgBook.Text, mtgPage.Text, mtgAddress.Text, satisfied, pdf, mtgBookSatisfied.Text, mtgPageSatisfied.Text, comboBox1.Text);
			}

			mtgAttButton.Content = "ATTACHMENT (PDF)";


		}

		private void AddMtgs(int id, string name, string amount, string loannum, DateTime date1, string book, string page, string address, string deedbook, string attachment, string bookSatisfied, string pageSatisfied, string satisfied)
		{
			if (mtgList.Count <= 12)
			{
				// mtgClass mtg = new mtgClass() { ID = mtgOnPropertyID, Name = mtgName.Text, Amount = mtgAmount.Text, LoanNum = mtgLoanNum.Text, Date = mtgDate.SelectedDate.Value, Book = mtgBook.Text, Page = mtgPage.Text, Address = mtgAddress.Text, DeedBook = satisfied };
				mtgClass mtg = new mtgClass() { ID = id, Name = name, Amount = amount, LoanNum = loannum, Date = date1, Book = book, Page = page, Address = address, DeedBook = deedbook, Attachment = attachment, BookSatisfied = bookSatisfied, PageSatisfied = pageSatisfied, Satisfied = comboBox1.Text };

				mtgList.Add(mtg);
				List<mtgClass> mtgSortedList = mtgList.OrderBy(o => o.Date).ToList();

				mtgList.Clear();
				mtgDataGrid.Items.Clear();

				foreach (mtgClass mtgcls in mtgSortedList)
				{
					mtgList.Add(mtgcls);
					mtgDataGrid.Items.Add(mtgcls);
				}
				
				mtgLabelCount.Content = mtgList.Count() + " OF 13";
				mtgID++;
				ClearEntryMtg();
				mtgName.Focus();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}

		private void AddMtgs(mtgClass mtg)
		{
			if (mtgList.Count <= 12)
			{

				mtgList.Add(mtg);
				List<mtgClass> mtgSortedList = mtgList.OrderBy(o => o.Date).ToList();

				mtgList.Clear();
				mtgDataGrid.Items.Clear();

				foreach (mtgClass mtgcls in mtgSortedList)
				{
					mtgList.Add(mtgcls);
					mtgDataGrid.Items.Add(mtgcls);
				}
				mtgLabelCount.Content = mtgList.Count() + " OF 13";
				ClearEntryMtg();
				mtgName.Focus();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}

		private void AddConclusion(int id, string name, string amount, DateTime date, string attachment, bool increment)
		{
			if (conclusionList.Count <= 9)
			{
				conclusionClass conCl = new conclusionClass() { ID = id, Name = name, Amount = amount, Date = date, Attachment = attachment };
				conclusionList.Add(conCl);
				List<conclusionClass> conclSortedList = conclusionList.OrderBy(o => o.Date).ToList();
				conclusionList.Clear();
				conclusionDataGrid.Items.Clear();

				foreach (conclusionClass conclusionCl in conclSortedList)
				{
					conclusionList.Add(conclusionCl);
					conclusionDataGrid.Items.Add(conclusionCl);
				}

				conclusionLabelCount.Content = conclusionList.Count() + " OF 10";
				if (increment) 
					conclusionID++;
			}
			else
			{
				MessageBox.Show("Maximum entries in CONCLUSION reached!");
			}
		}

		private void AddConclusion(conclusionClass conCl)
		{
			if (conclusionList.Count <= 9)
			{
				//conclusionClass conCl = new conclusionClass() { ID = conclusionID, Name = name, Amount = amount, Date = date };
				conclusionList.Add(conCl);
				List<conclusionClass> conclSortedList = conclusionList.OrderBy(o => o.Date).ToList();
				conclusionList.Clear();
				conclusionDataGrid.Items.Clear();
				
				foreach (conclusionClass conclusionCl in conclSortedList)
				{
					conclusionList.Add(conclusionCl);
					conclusionDataGrid.Items.Add(conclusionCl);
				}
				
				conclusionLabelCount.Content = conclusionList.Count() + " OF 10";
			}
			else
			{
				MessageBox.Show("Maximum entries in CONCLUSION reached!");
			}
		}

		private void button6_Click(object sender, RoutedEventArgs e)
		{
			string pdf = judAttButton.Content.ToString().Replace("ATTACHMENT (PDF)", "");
			ComboBoxItem cbItem = (ComboBoxItem)comboBox1.SelectedItem;

			if (isJudLoaded)
			{

				DeleteConclusion(selectedID);
				AddConclusion(selectedID, judName.Text, judCurrAmount.Text, judDateRec.SelectedDate.Value, pdf, false);

				DeleteJud(selectedID);

				AddJudgCases(selectedID, judCaseNum.Text, judDateRec.SelectedDate.Value, judOrigAmount.Text, judCurrAmount.Text, judName.Text + " " + judContact.Text, pdf, judName.Text, judContact.Text);

				isJudLoaded = false;
			}
			else
			{
				while ((judList.Find(x => x.ID == judgementID)) != null)
				{
					judgementID++;
				}
				
				AddConclusion(judgementID, judName.Text, judCurrAmount.Text, judDateRec.SelectedDate.Value, pdf, false);
				AddJudgCases(judgementID, judCaseNum.Text, judDateRec.SelectedDate.Value, judOrigAmount.Text, judCurrAmount.Text, judName.Text + " " + judContact.Text, pdf, judName.Text, judContact.Text);
			}

			ClearEntryJud();
			judAttButton.Content = "ATTACHMENT (PDF)";
		}

		private void AddJudgCases(int id, string judcasenum, DateTime recorded, string origamount, string curamount, string contact, string attachment, string name, string contact2)
		{
			if (judList.Count <= 11)
			{
				//judgementClass judLien = new judgementClass() { ID = judgementID, CaseNum = judCaseNum.Text, DateRec = judDateRec.SelectedDate.Value, OriginalAmount = judOrigAmount.Text, CurrentAmount = judCurrAmount.Text, Contact = judName.Text + " " + judContact.Text };
				judgementClass judLien = new judgementClass() { ID = id, CaseNum = judcasenum, DateRec = recorded, OriginalAmount = origamount, CurrentAmount = curamount, Contact = contact, Attachment = attachment, Name = name, Contact2 = contact2 };
				judList.Add(judLien);
				
				List<judgementClass> judSortedList = judList.OrderBy(o => o.DateRec).ToList();
				judList.Clear();
				judDataGrid.Items.Clear();
				
				foreach (judgementClass judcls in judSortedList)
				{
					judList.Add(judcls);
					judDataGrid.Items.Add(judcls);
				}

				judgementLabelCount.Content = judList.Count() + " OF 12";
				judgementID++;
				ClearEntryJud();
				judCaseNum.Focus();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}
		private void AddJudgCases(judgementClass judLien)
		{
			if (judList.Count <= 11)
			{
				judList.Add(judLien);
				List<judgementClass> judSortedList = judList.OrderBy(o => o.DateRec).ToList();
				judList.Clear();
				judDataGrid.Items.Clear();
				
				foreach (judgementClass judcls in judSortedList)
				{
					judList.Add(judcls);
					judDataGrid.Items.Add(judcls);
				}
				judgementLabelCount.Content = judList.Count() + " OF 12";

				ClearEntryJud();
				judCaseNum.Focus();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}

		public void ClearEntryMtg()
		{
			mtgName.Text = "";
			mtgAmount.Text = "";
			mtgLoanNum.Text = "";
			mtgDate.SelectedDate = DateTime.Now;

			mtgBook.Text = "";
			mtgPage.Text = "";
			mtgAddress.Text = "";

			mtgBookSatisfied.Text = "";
			mtgPageSatisfied.Text = "";
			comboBox1.SelectedIndex = 0;


			satisfied = "";
		}

		public void ClearEntryJud()
		{
			judCaseNum.Text = "";
			judDateRec.SelectedDate = DateTime.Now;
			judOrigAmount.Text = "";
			judCurrAmount.Text = "";
			judContact.Text = "";
			judName.Text = "";

		}

		public void ClearEntryClaimant()
		{

			tbClaimantName.Text = "";
			tbClaimantPhone.Text = "";
			tbClaimantAddress.Text = "";
			tbClaimantGuess.Text = "";

		}


		private void ClearEntryConclusion()
		{
			tbConclusionName.Text = "";
			tbConclusionAmount.Text = "";
			conclusionDate.SelectedDate = DateTime.Now;

		}

		public class mtgClass
		{
			public int ID { get; set; }
			public string Name { get; set; }
			public string Amount { get; set; }
			public string LoanNum { get; set; }
			public DateTime Date { get; set; }
			public string Recorded { get; set; }
			public string Book { get; set; }
			public string Page { get; set; }
			public string Address { get; set; }
			public string DeedBook { get; set; }

			public string Satisfied { get; set; }
			public string BookSatisfied { get; set; }

			public string PageSatisfied { get; set; }
			public string Attachment { get; set; }

		}

		public class judgementClass
		{
			public int ID { get; set; }
			public string CaseNum { get; set; }
			public DateTime DateRec { get; set; }
			public string OriginalAmount { get; set; }
			public string CurrentAmount { get; set; }

			public string Name { get; set; }
			public string Contact2 { get; set; }

			public string Address { get; set; }
			public string Contact { get; set; }
			public string Attachment { get; set; }

		}

		public class conclusionClass
		{
			public int ID { get; set; }
			public string Name { get; set; }
			public string Amount { get; set; }
			public DateTime Date { get; set; }
			public string Attachment { get; set; }
		}

		public class claimantsClass
		{
			public int ID { get; set; }
			public string ClaimantName { get; set; }
			public string ClaimantPhone { get; set; }
			public string ClaimantAddress { get; set; }
			public string ClaimantGuess { get; set; }

		}

		private void button5_Click(object sender, RoutedEventArgs e)
		{
			expander2.IsExpanded = false;
			expander3.IsExpanded = true;
			judCaseNum.Focus();
		}

		private void dataGrid1_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Delete)
			{
				if (mtgDataGrid.SelectedIndex >= 0)
				{
					MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
					if (mbres == MessageBoxResult.Yes)
					{
						mtgClass mtgTemp = (mtgClass)mtgDataGrid.SelectedItem;

						DeleteMtg(mtgTemp.ID);
						DeleteConclusion(mtgTemp.ID);
					}
				}
			}
		}

		private void DeleteMtg(int id)
		{
			mtgList.Remove(mtgList.Find(x => x.ID == id));
			List<mtgClass> mtgSortedList = mtgList.OrderBy(o => o.Date).ToList();
			mtgDataGrid.Items.Clear();
			mtgList.Clear();
			foreach (mtgClass mtgcls in mtgSortedList)
			{
				mtgList.Add(mtgcls);
				mtgDataGrid.Items.Add(mtgcls);
			}
			mtgLabelCount.Content = mtgList.Count() + " OF 13";
		}


		private void DeleteJud(int id)
		{
			judList.Remove(judList.Find(x => x.ID == id));
			List<judgementClass> judSortedList = judList.OrderBy(o => o.DateRec).ToList();
			judDataGrid.Items.Clear();
			judList.Clear();
			foreach (judgementClass judcls in judSortedList)
			{
				judList.Add(judcls);
				judDataGrid.Items.Add(judcls);
			}
			judgementLabelCount.Content = judList.Count() + " OF 12";
		}
		private void button7_Click(object sender, RoutedEventArgs e)
		{
			expander3.IsExpanded = false;
			expander4.IsExpanded = true;

		}
		private void judDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Delete)
			{
				if (judDataGrid.SelectedIndex >= 0)
				{
					MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
					if (mbres == MessageBoxResult.Yes)
					{
						judgementClass judTemp = (judgementClass)judDataGrid.SelectedItem;
						DeleteJud(judTemp.ID);
						DeleteConclusion(judTemp.ID);
					}
				}
			}
		}
		private void button9_Click(object sender, RoutedEventArgs e)
		{
			if (isClaimantLoaded)
			{
				DeleteClaimant(selectedID);
				AddClaimants(selectedID, tbClaimantName.Text, tbClaimantPhone.Text, tbClaimantAddress.Text, tbClaimantGuess.Text);
				isClaimantLoaded = false;
			}

			else
			{
				AddClaimants(claimantID, tbClaimantName.Text, tbClaimantPhone.Text, tbClaimantAddress.Text, tbClaimantGuess.Text);
			}
		}

		private void AddClaimants(int id, string name, string phone, string address, string guess)
		{
			if (claimantList.Count <= 6)
			{
				//  claimantsClass claimcls = new claimantsClass() { ClaimantName = tbClaimantName.Text, ClaimantPhone = tbClaimantPhone.Text, ClaimantAddress = tbClaimantAddress.Text, ClaimantGuess = tbClaimantGuess.Text };

				claimantsClass claimcls = new claimantsClass() { ID = id, ClaimantName = name, ClaimantPhone = phone, ClaimantAddress = address, ClaimantGuess = guess };
				claimantList.Add(claimcls);

				claimantsDataGrid.Items.Clear();

				foreach (claimantsClass cls in claimantList)
				{
					claimantsDataGrid.Items.Add(cls);
				}

				claimantID++;
				claimantsLabelCount.Content = claimantList.Count.ToString() + " OF 7";

				ClearEntryClaimant();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}

		private void AddClaimants(claimantsClass claimcls)
		{
			if (claimantList.Count <= 6)
			{
				claimantList.Add(claimcls);
				claimantsDataGrid.Items.Clear();
				foreach (claimantsClass cls in claimantList)
				{
					claimantsDataGrid.Items.Add(cls);
				}

				claimantID++;
				claimantsLabelCount.Content = claimantList.Count.ToString() + " OF 7";

				ClearEntryClaimant();
			}
			else
			{
				MessageBox.Show("Maximum entries reached!");
			}
		}
		private void expander2_Loaded(object sender, RoutedEventArgs e)
		{
			mtgName.Focus();
		}

		private void Submit(bool saveAndCreate)
		{
			bool extended = false;
			xlTemplatePath = Properties.Settings.Default.pathTemplate;
			if (File.Exists(xlTemplatePath) && Directory.Exists(Properties.Settings.Default.pathSaveFolder))
			{
				if (tbName.Text.Contains(','))
				{
					string state = tbState.Text.ToUpper();
					string county = tbCounty.Text.ToUpper();



					if (isMtgForm)
						documentDir = Properties.Settings.Default.pathSaveFolder + "\\" + state + "\\" + county + "\\MORTGAGE_FORMS\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + "\\";
					if (isDeedsForm)
						documentDir = Properties.Settings.Default.pathSaveFolder + "\\" + state + "\\" + county + "\\DEEDS_FORMS\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + "\\";
					else
						documentDir = Properties.Settings.Default.pathSaveFolder + "\\" + state + "\\" + county + "\\TAX_FORMS\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + "\\";

					pdfSavePath = documentDir + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + ".pdf";
					xlSavePath = documentDir + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + ".xlsm";

					if (File.Exists(pdfSavePath))
					{
						MessageBoxResult mbrs = MessageBox.Show("The file " + System.IO.Path.GetFileName(pdfSavePath) + " already exists! OVERWRITE?", "File Exists", MessageBoxButton.YesNo);

						if (mbrs == MessageBoxResult.Yes)
						{

							try
							{
								WriteToExcel(xlTemplatePath, xlSavePath, pdfSavePath, extended, saveAndCreate);
							}
							catch (Exception Ex)
							{
								MessageBox.Show("Error 1: " + Ex.Message);
							}

							try
							{
								XmlSave();
							}
							catch (Exception Ex)
							{
								MessageBox.Show("Error 2: " + Ex.Message);
							}


						}
					}

					else
					{

						try
						{
							WriteToExcel(xlTemplatePath, xlSavePath, pdfSavePath, extended, saveAndCreate);
						}
						catch (Exception Ex)
						{
							MessageBox.Show("Error 1: " + Ex.Message);
						}

						try
						{
							XmlSave();
						}
						catch (Exception Ex)
						{
							MessageBox.Show("Error 2: " + Ex.Message);
						}

					}

				}

				else
				{
					MessageBox.Show("Missing comma delimiter in [Names of Persons] field");
				}

			}
			else
			{
				MessageBox.Show("Paths not set! Set Paths and try again.");
			}
		}


		private void WriteToExcel(string templatePath, string xlPath, string pdfPath, bool extended, bool saveAndCreate)
		{


			if (!Directory.Exists(documentDir))
			{
				Directory.CreateDirectory(documentDir);
				Directory.CreateDirectory(documentDir + "\\files");
			}

			File.Copy(templatePath, xlPath, true);

			if (textBox30.Text.Length > 0)
				pdfCheck1 = true;
			if (textBox31.Text.Length > 0)
				pdfCheck2 = true;
			if (textBox32.Text.Length > 0)
				pdfCheck3 = true;
			if (textBox33.Text.Length > 0)
				pdfCheck5 = true;
			if (textBox34.Text.Length > 0)
				pdfCheck6 = true;

			if (mtgList.Count > 9)
				mtgExtended = true;

			if (judList.Count > 7)
				judgmentExtended = true;

			if (conclusionList.Count > 4)
				conclusionExtended = true;

			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWb = xlApp.Workbooks.Open(xlPath, false, false);
			Excel.Worksheet xlWs = xlWb.Worksheets["data"];
			Excel.Worksheet xlTax = xlWb.Worksheets["tax"];
			Excel.Worksheet xlEstateData = xlWb.Worksheets["estates_data"];
			Excel.Worksheet xlEstate = xlWb.Worksheets["estates"];

			//MTG FORCLOSURE PAGE
			xlWs.get_Range("A2").Value2 = 1;
			xlWs.get_Range("B2").Value2 = tbName.Text;
			xlWs.get_Range("C2").Value2 = tbDeedHolder.Text;
			xlWs.get_Range("D2").Value2 = tbFile.Text;

			if (tbSurplus.Text.Length > 0)
				xlWs.get_Range("E2").Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbSurplus.Text));

			xlWs.get_Range("F2").Value2 = tbCounty.Text;
			xlWs.get_Range("G2").Value2 = tbState.Text;
			xlWs.get_Range("H2").Value2 = dateReviewed.SelectedDate;
			xlWs.get_Range("I2").Value2 = tbResearcher.Text;
			xlWs.get_Range("J2").Value2 = tbHeld.Text;
			xlWs.get_Range("K2").Value2 = tbVerify.Text;
			xlWs.get_Range("L2").Value2 = tbAddress.Text;
			xlWs.get_Range("M2").Value2 = dateRecorded.SelectedDate;
			xlWs.get_Range("N2").Value2 = tbBook.Text;
			xlWs.get_Range("O2").Value2 = tbPage.Text;
			xlWs.get_Range("P2").Value2 = dateForclosed.SelectedDate;

			if (tbNote.Text.Length > 0)
				xlWs.get_Range("Q2").Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbNote.Text));

			xlWs.get_Range("R2").Value2 = tbBeneficiary.Text;
			//MTG FORCLOSURE PAGE

			//MTG HISTORY PAGE
			//using column id instead of column name - starts with col 19
			int mtgColCount = 19;
			foreach (mtgClass mtg in mtgList)
			{
				xlWs.Cells[2, mtgColCount].Value2 = mtg.Name;

				if (mtg.Amount != null)
				{
					if (mtg.Amount.Length > 0)
						xlWs.Cells[2, mtgColCount + 1].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(mtg.Amount));
				}
				xlWs.Cells[2, mtgColCount + 2].Value2 = mtg.Date;
				xlWs.Cells[2, mtgColCount + 3].Value2 = mtg.Date;
				xlWs.Cells[2, mtgColCount + 4].Value2 = mtg.Book;
				xlWs.Cells[2, mtgColCount + 5].Value2 = mtg.Page;


				xlWs.Cells[2, mtgColCount + 6].Value2 = mtg.LoanNum;

				xlWs.Cells[2, mtgColCount + 7].Value2 = mtg.Address;
				xlWs.Cells[2, mtgColCount + 8].Value2 = mtg.DeedBook;
				xlWs.Cells[2, mtgColCount + 9].Value2 = mtg.Attachment;

				mtgColCount += 10;
			}
			//JUDGMENT HISTORY PAGE

			//using column id instead of column name - starts with col 19
			int judColCount = 149;

			foreach (judgementClass jud in judList)
			{
				xlWs.Cells[2, judColCount].Value2 = jud.CaseNum;
				xlWs.Cells[2, judColCount + 1].Value2 = jud.DateRec;

				if (jud.OriginalAmount != null)
				{
					if (jud.OriginalAmount.Length > 0)
						xlWs.Cells[2, judColCount + 2].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(jud.OriginalAmount));
				}

				if (jud.CurrentAmount != null)
				{
					if (jud.CurrentAmount.Length > 0)
						xlWs.Cells[2, judColCount + 3].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(jud.CurrentAmount));
				}

				xlWs.Cells[2, judColCount + 4].Value2 = jud.Contact;
				xlWs.Cells[2, judColCount + 5].Value2 = jud.Attachment;

				judColCount += 6;
			}
			//JUDGMENT HISTORY PAGE


			//using column id instead of column name - starts with col 19
			int dueCount = 220;


			foreach (conclusionClass conccls in conclusionList)
			{
				xlWs.Cells[2, dueCount].Value2 = conccls.Name;

				if (conccls.Amount != null)
				{
					if (conccls.Amount.Length > 0)
						xlWs.Cells[2, dueCount + 1].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(conccls.Amount));
				}
				xlWs.Cells[2, dueCount + 2].Value2 = conccls.Date;
				xlWs.Cells[2, dueCount + 3].Value2 = conccls.Attachment;

				dueCount += 4;

			}

			int claimantCount = 260;


			foreach (claimantsClass claimantcls in claimantList)
			{
				xlWs.Cells[2, claimantCount].Value2 = claimantcls.ClaimantName;
				xlWs.Cells[2, claimantCount + 1].Value2 = claimantcls.ClaimantPhone;
				xlWs.Cells[2, claimantCount + 2].Value2 = claimantcls.ClaimantAddress;
				xlWs.Cells[2, claimantCount + 3].Value2 = claimantcls.ClaimantGuess;

				claimantCount += 4;

			}

			//JUDGMENT HISTORY PAGE

			xlEstateData.Cells[2, 1].Value2 = tbEstatesNameDeceased.Text;
			xlEstateData.Cells[2, 2].Value2 = tbEstatesFileNumber.Text;

			if (tbEstatesFunds.Text.Length > 0)
				xlEstateData.Cells[2, 3].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbEstatesFunds.Text));


			xlEstateData.Cells[2, 4].Value2 = cbEstatesCounty.Text;
			xlEstateData.Cells[2, 5].Value2 = cbEstatesState.Text;
			xlEstateData.Cells[2, 6].Value2 = dtEstateDateReviewed.SelectedDate.Value;
			xlEstateData.Cells[2, 7].Value2 = tbEstatesResearcher.Text;
			xlEstateData.Cells[2, 8].Value2 = tbEstatePartOf.Text;
			xlEstateData.Cells[2, 9].Value2 = tbEstatesHow.Text;

			xlEstateData.Cells[2, 10].Value2 = tbEstatesName1.Text;

			if (tbEstatesAmount1.Text.Length > 0)
				xlEstateData.Cells[2, 11].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbEstatesAmount1.Text));

			xlEstateData.Cells[2, 12].Value2 = tbEstatesName2.Text;

			if (tbEstatesAmount2.Text.Length > 0)
				xlEstateData.Cells[2, 13].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbEstatesAmount2.Text));

			xlEstateData.Cells[2, 14].Value2 = tbEstatesName3.Text;

			if (tbEstatesAmount3.Text.Length > 0)
				xlEstateData.Cells[2, 15].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbEstatesAmount3.Text));

			xlEstateData.Cells[2, 16].Value2 = tbEstatesName4.Text;

			if (tbEstatesAmount4.Text.Length > 0)
				xlEstateData.Cells[2, 17].Value2 = string.Format("{0:#,#.00}", Convert.ToDouble(tbEstatesAmount4.Text));


			if (tbEstatesDeathAtt.Text.Length > 0)
				xlEstate.Range["M17"].Value2 = "TRUE";
			else
				xlEstate.Range["M17"].Value2 = "FALSE";



			if (tbEstatesWillAtt.Text.Length > 0)
				xlEstate.Range["M18"].Value2 = "TRUE";
			else
				xlEstate.Range["M18"].Value2 = "FALSE";

			if (listEstateAnyDocs.Items.Count > 0)
				xlEstate.Range["M20"].Value2 = "TRUE";
			else
				xlEstate.Range["M20"].Value2 = "FALSE";



			if (tbEstatesNotifAtt.Text.Length > 0)
				xlEstate.Range["M21"].Value2 = "TRUE";
			else
				xlEstate.Range["M21"].Value2 = "FALSE";


			if (listEstateAnyOtherDocs.Items.Count > 0)
				xlEstate.Range["M22"].Value2 = "TRUE";
			else
				xlEstate.Range["M22"].Value2 = "FALSE";



			xlWb.Save();

			xlApp.DisplayAlerts = false;




			Excel.Worksheet xlWsPdf = xlWb.Worksheets["PDFEXTENDED"];

			xlWsPdf.Range["M27"].Value2 = pdfCheck1.ToString();
			xlWsPdf.Range["M28"].Value2 = pdfCheck2.ToString();
			xlWsPdf.Range["M30"].Value2 = pdfCheck3.ToString();

			xlTax.Range["M27"].Value2 = pdfCheck1.ToString();
			xlTax.Range["M28"].Value2 = pdfCheck2.ToString();
			xlTax.Range["M30"].Value2 = pdfCheck3.ToString();

			foreach (mtgClass mtg in mtgList)
			{
				if (mtg.Attachment != null)
				{
					if (mtg.Attachment.Length > 0)
					{
						pdfCheck4 = true;
					}
				}
			}

			foreach (judgementClass judCls in judList)
			{
				if (judCls.Attachment != null)
				{
					if (judCls.Attachment.Length > 0)
					{
						pdfCheck4 = true;
					}
				}
			}

			xlWsPdf.Range["M31"].Value2 = pdfCheck4.ToString();
			xlTax.Range["M31"].Value2 = pdfCheck4.ToString();



			firstPagesMtgPDF = pdfPath.Replace(".pdf", "firstmtg.pdf");
			firstPagesTaxPDF = pdfPath.Replace(".pdf", "firsttax.pdf");
			mtgPagesPDF = pdfPath.Replace(".pdf", "mtg.pdf");
			judgmentpagesPDF = pdfPath.Replace(".pdf", "jud.pdf");
			conclusionPagesPDF = pdfPath.Replace(".pdf", "concl.pdf");
			estatePagePdf = pdfPath.Replace(".pdf", "estate.pdf");

			if (isMtgForm)
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, firstPagesMtgPDF, IgnorePrintAreas: false, From: 1, To: 3);
			if (isDeedsForm)
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, firstPagesMtgPDF, IgnorePrintAreas: false, From: 1, To: 3);
			else
			{
				xlTax.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, firstPagesTaxPDF, IgnorePrintAreas: false, From: 1, To: 1);
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, firstPagesMtgPDF, IgnorePrintAreas: false, From: 2, To: 3);
			}

			if (mtgExtended)
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, mtgPagesPDF, IgnorePrintAreas: false, From: 4, To: 4);

			if (judgmentExtended)
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, judgmentpagesPDF, IgnorePrintAreas: false, From: 5, To: 6);
			else
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, judgmentpagesPDF, IgnorePrintAreas: false, From: 5, To: 5);

			if (conclusionExtended)
				xlWsPdf.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, conclusionPagesPDF, IgnorePrintAreas: false, From: 7, To: 7);
			else
			{
				Excel.Worksheet xlWsSimplified = xlWb.Worksheets["PDFSIMPLIFIED"];
				xlWsSimplified.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, conclusionPagesPDF, IgnorePrintAreas: false, From: 5, To: 5);
			}


			if (cbAddEstates.IsChecked.Value == true)
			{
				xlEstate.UsedRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, estatePagePdf, IgnorePrintAreas: false, From: 1, To: 1);
			}



			xlWb.Close(true);
			xlApp.DisplayAlerts = true;

			xlApp.Quit();



			docs = new List<string>();

			if (isMtgForm)
			{
				docs.Add(firstPagesMtgPDF);
			}
			if (isDeedsForm)
				docs.Add(firstPagesMtgPDF);
			else
			{
				docs.Add(firstPagesTaxPDF);
				docs.Add(firstPagesMtgPDF);
			}


			if (mtgExtended)
				docs.Add(mtgPagesPDF);

			docs.Add(judgmentpagesPDF);

			docs.Add(conclusionPagesPDF);

			if (cbAddEstates.IsChecked.Value == true)
			{
				docs.Add(estatePagePdf);
			}

			if (textBox30.Text.Length > 4)
				p1.pdfForclosureOrder = SavePDF(textBox30.Text, documentDir + "files\\" + System.IO.Path.GetFileName(textBox30.Text));

			if (textBox31.Text.Length > 4)
				p1.pdfDisbursement = SavePDF(textBox31.Text, documentDir + "files\\" + System.IO.Path.GetFileName(textBox31.Text));

			if (textBox32.Text.Length > 4)
				p1.pdfNotification = SavePDF(textBox32.Text, documentDir + "files\\" + System.IO.Path.GetFileName(textBox32.Text));



			if (textBox33.Text.Length > 4)
				p1.pdfDeed = SavePDF(textBox33.Text, documentDir + "files\\" + System.IO.Path.GetFileName(textBox33.Text));

			if (textBox34.Text.Length > 4)
				p1.pdfSheriffDeed = SavePDF(textBox34.Text, documentDir + "files\\" + System.IO.Path.GetFileName(textBox34.Text));




			List<conclusionClass> conclSortedList = conclusionList.OrderBy(o => o.Date).ToList();

			foreach (conclusionClass conCls in conclSortedList)
			{
				if (conCls.Attachment.Length > 4)
				{
					conCls.Attachment = SavePDF(conCls.Attachment, documentDir + "files\\" + System.IO.Path.GetFileName(conCls.Attachment));
					try
					{
						mtgList.Find(x => x.ID == conCls.ID).Attachment = conCls.Attachment;
					}
					catch { }

					try
					{
						judList.Find(x => x.ID == conCls.ID).Attachment = conCls.Attachment;
					}
					catch { }
				}
			}

			if (tbEstatesDeathAtt.Text.Length > 0)
			{
				p1.pdfEstateDeathCert = SavePDF(tbEstatesDeathAtt.Text, documentDir + "files\\" + System.IO.Path.GetFileName(tbEstatesDeathAtt.Text));
			}

			if (tbEstatesWillAtt.Text.Length > 0)
			{

				p1.pdfEstateWill = SavePDF(tbEstatesWillAtt.Text, documentDir + "files\\" + System.IO.Path.GetFileName(tbEstatesWillAtt.Text));

			}


			anyDocsList.Clear();

			foreach (string item in listEstateAnyDocs.Items)
			{
				anyDocsList.Add(SavePDF(item, documentDir + "files\\" + System.IO.Path.GetFileName(item)));
			}

			if (tbEstatesNotifAtt.Text.Length > 0)
			{
				p1.pdfEstateNotif = SavePDF(tbEstatesNotifAtt.Text, documentDir + "files\\" + System.IO.Path.GetFileName(tbEstatesNotifAtt.Text));
			}

			anyOtherDocsList.Clear();

			foreach (string item in listEstateAnyOtherDocs.Items)
			{
				anyOtherDocsList.Add(SavePDF(item, documentDir + "files\\" + System.IO.Path.GetFileName(item)));
			}


			tempMisc.Clear();

			foreach (string misc in listBoxPDF.Items)
			{
				tempMisc.Add(SavePDF(misc, documentDir + "files\\" + System.IO.Path.GetFileName(misc)));
			}


			if (docs.Count > 1)
			{
				CombineMultiplePDFs(docs.ToArray(), pdfPath);
			}

			try
			{
				File.Delete(pdfPath.Replace(".pdf", "firstmtg.pdf"));
				File.Delete(pdfPath.Replace(".pdf", "firsttax.pdf"));
			}
			catch{}

			try
			{
				File.Delete(conclusionPagesPDF);
			}
			catch{}
			try
			{
				File.Delete(judgmentpagesPDF);
			}
			catch{}

			try
			{
				File.Delete(estatePagePdf);
			}
			catch{}

			if (mtgExtended)
				File.Delete(mtgPagesPDF);

			if (saveAndCreate)
				System.Diagnostics.Process.Start(pdfPath);
		}



		private string SavePDF(string sourceFile, string destinationFile)
		{

			int id = 1;
			sourceFile = sourceFile.ToLower();
			destinationFile = destinationFile.ToLower();

			if (sourceFile != destinationFile)
			{
				while (File.Exists(destinationFile))
				{
					if (id > 1)
						destinationFile = destinationFile.Substring(0, destinationFile.Length - 5) + id.ToString() + ".pdf";
					else
						destinationFile = destinationFile.Substring(0, destinationFile.Length - 4) + id.ToString() + ".pdf";
					id++;
				}

				File.Copy(sourceFile, destinationFile, true);
			}


			docs.Add(destinationFile);

			return destinationFile;
		}
		private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{

			ComboBoxItem cbItem = (ComboBoxItem)comboBox1.SelectedItem;
			if (cbItem != null)
			{
				if (cbItem.Content.ToString() != "-")
				{
					satisfied = cbItem.Content + " - BOOK: " + mtgBookSatisfied.Text.ToUpper() + "; PAGE: " + mtgPageSatisfied.Text.ToUpper();

					mtgBookSatisfied.Visibility = System.Windows.Visibility.Visible;
					mtgPageSatisfied.Visibility = System.Windows.Visibility.Visible;

					lblBook.Visibility = System.Windows.Visibility.Visible;
					lblPage.Visibility = System.Windows.Visibility.Visible;
				}
				else
				{
					mtgBookSatisfied.Visibility = System.Windows.Visibility.Hidden;
					mtgPageSatisfied.Visibility = System.Windows.Visibility.Hidden;

					lblBook.Visibility = System.Windows.Visibility.Hidden;
					lblPage.Visibility = System.Windows.Visibility.Hidden;

					mtgAttButton.Visibility = System.Windows.Visibility.Visible;
					satisfied = "";
				}
			}
		}

		private void mtgBookSatisfied_TextChanged(object sender, TextChangedEventArgs e)
		{
			ComboBoxItem cbItem = (ComboBoxItem)comboBox1.SelectedItem;
			if (cbItem != null)
				satisfied = cbItem.Content + " - BOOK: " + mtgBookSatisfied.Text.ToUpper() + "; PAGE: " + mtgPageSatisfied.Text.ToUpper();

		}

		private void mtgBookPage_TextChanged(object sender, TextChangedEventArgs e)
		{
			ComboBoxItem cbItem = (ComboBoxItem)comboBox1.SelectedItem;
			if (cbItem != null)
				satisfied = cbItem.Content + " - BOOK: " + mtgBookSatisfied.Text.ToUpper() + "; PAGE: " + mtgPageSatisfied.Text.ToUpper();
		}

		private void claimantsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Delete)
			{
				if (claimantsDataGrid.SelectedIndex >= 0)
				{
					MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
					if (mbres == MessageBoxResult.Yes)
					{
						claimantsClass clCls = (claimantsClass)claimantsDataGrid.SelectedItem;
						DeleteClaimant(clCls.ID);
					}
				}
			}
		}

		private void DeleteClaimant(int id)
		{
			claimantsClass clTemp = claimantList.Find(x => x.ID == id);

			claimantsDataGrid.Items.Remove(clTemp);
			claimantList.Remove(clTemp);

			claimantsLabelCount.Content = claimantList.Count.ToString() + " OF 7";

		}
		private void button8_Click_1(object sender, RoutedEventArgs e)
		{
			if (cbAddEstates.IsChecked.Value == true)
			{
				expander4.IsExpanded = false;
				expander8.IsExpanded = true;
			}
			else
			{
				expander4.IsExpanded = false;
				expander6.IsExpanded = true;
			}
		}

		private void conclusionDataGrid_PreviewKeyDown_2(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Delete)
			{
				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);

				if (mbres == MessageBoxResult.Yes)
				{
					conclusionClass conclTemp = (conclusionClass)conclusionDataGrid.SelectedItem;
					DeleteConclusion(conclTemp.ID);
				}
			}
		}

		private void button12_Click(object sender, RoutedEventArgs e)
		{
			if (mtgDataGrid.SelectedIndex >= 0)
			{
				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
				if (mbres == MessageBoxResult.Yes)
				{
					mtgClass mtgTemp = (mtgClass)mtgDataGrid.SelectedItem;
					DeleteMtg(mtgTemp.ID);
					DeleteConclusion(mtgTemp.ID);
				}
			}

		}

		private void button14_Click(object sender, RoutedEventArgs e)
		{
			if (conclusionDataGrid.SelectedIndex >= 0)
			{
				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);

				if (mbres == MessageBoxResult.Yes)
				{
					conclusionClass conclTemp = (conclusionClass)conclusionDataGrid.SelectedItem;
					DeleteConclusion(conclTemp.ID);
				}
			}
		}

		private void DeleteConclusion(int id)
		{
			conclusionList.Remove(conclusionList.Find(x => x.ID == id));
			List<conclusionClass> conclSortedList = conclusionList.OrderBy(o => o.Date).ToList();
			conclusionList.Clear();
			conclusionDataGrid.Items.Clear();
			foreach (conclusionClass conclusionCl in conclSortedList)
			{
				conclusionList.Add(conclusionCl);
				conclusionDataGrid.Items.Add(conclusionCl);
			}
			conclusionID = conclusionList.Count();
			conclusionLabelCount.Content = (conclusionID) + " OF 10";
		}

		private void button13_Click(object sender, RoutedEventArgs e)
		{
			if (judDataGrid.SelectedIndex >= 0)
			{
				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
				if (mbres == MessageBoxResult.Yes)
				{
					judgementClass judTemp = (judgementClass)judDataGrid.SelectedItem;
					DeleteJud(judTemp.ID);
					DeleteConclusion(judTemp.ID);
				}
			}
		}

		private void button16_Click(object sender, RoutedEventArgs e)
		{
			string pdf = conclusionAttButton.Content.ToString().Replace("ATTACHMENT (PDF)", "");

			if (isConcLoaded)
			{
				DeleteConclusion(selectedID);
				AddConclusion(selectedID, tbConclusionName.Text, tbConclusionAmount.Text, conclusionDate.SelectedDate.Value, pdf, true);
				isConcLoaded = false;
			}
			else
			{
				while (conclusionList.Find(x => x.ID == conclusionID) != null)
				{
					conclusionID++;
				}
				AddConclusion(conclusionID, tbConclusionName.Text, tbConclusionAmount.Text, conclusionDate.SelectedDate.Value, pdf, true);
			}
			ClearEntryConclusion();
			conclusionAttButton.Content = "ATTACHMENT (PDF)";
		}



		private void button10_Click_1(object sender, RoutedEventArgs e)
		{
			expander6.IsExpanded = false;
			expander5.IsExpanded = true;
		}

		private void button15_Click_1(object sender, RoutedEventArgs e)
		{
			if (claimantsDataGrid.SelectedIndex >= 0)
			{

				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete this row?", "!", MessageBoxButton.YesNo);
				if (mbres == MessageBoxResult.Yes)
				{

					claimantsClass clCls = (claimantsClass)claimantsDataGrid.SelectedItem;

					DeleteClaimant(clCls.ID);
				}
			}
		}

		private void tbName_LostFocus(object sender, RoutedEventArgs e)
		{

			if (tbName.Text.Length > 0)
			{
				char[] toReplace = { '\\', '/', ':', '*', '?', '\"', '<', '>', '|' };
				foreach (char chr in toReplace)
					tbName.Text = tbName.Text.Replace(chr.ToString(), "");
			}
		}

		private void tbName_TextChanged(object sender, TextChangedEventArgs e)
		{
			try
			{
				tbEstatesNameDeceased.Text = tbName.Text.Substring(0, tbName.Text.IndexOf(","));
			}
			catch{}
		}

		private void tbSurplus_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void tbNote_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void mtgAmount_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void judOrigAmount_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void judCurrAmount_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void tbConclusionAmount_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.OemComma)
			{
				e.Handled = true;
			}

			else
			{
				e.Handled = false;
			}
		}

		private void button18_Click(object sender, RoutedEventArgs e)
		{

			if (tbState.Text == "" || tbCounty.Text == "")
			{
				MessageBox.Show("Please fill-in State and County");
			}
			else
			{
				Submit(true);
			}

		}
		private void mtgForm_Click(object sender, RoutedEventArgs e)
		{
			isMtgForm = true;
			isTaxForm = false;
			MainWindow mw = new MainWindow();

			mw.Show();

			mw.ScrollViewer.Visibility = System.Windows.Visibility.Visible;
			this.Close();

		}


		private void taxForm_Click(object sender, RoutedEventArgs e)
		{
			isMtgForm = false;
			isTaxForm = true;
			MainWindow mw = new MainWindow();

			mw.Show();

			mw.ScrollViewer.Visibility = System.Windows.Visibility.Visible;
			this.Close();
		}

		private void Modify_Click(object sender, RoutedEventArgs e)
		{
			OpenFile op = new OpenFile();
			op.ShowDialog();
			if (FileToOpen != "")
			{
				MainWindow mw = new MainWindow();
				this.Close();
				mw.Show();

				mw.ScrollViewer.Visibility = Visibility.Visible;
				mw.ReloadData(FileToOpen);
			}
		}
		private void ReloadData(string xmlPath)
		{
			p1 = new project1();
			XmlSerializer serializer = new XmlSerializer(p1.GetType());
			FileStream fs = new FileStream(xmlPath, FileMode.Open);

			p1 = (project1)serializer.Deserialize(fs);

			textBox30.Text = p1.pdfForclosureOrder;
			textBox31.Text = p1.pdfDisbursement;
			textBox32.Text = p1.pdfNotification;
			textBox33.Text = p1.pdfDeed;
			textBox34.Text = p1.pdfSheriffDeed;

			foreach (string misc in p1.miscPdfs)
			{
				listBoxPDF.Items.Add(misc);
			}

			tbName.Text = p1.mtgName;
			tbDeedHolder.Text = p1.deedHolder;
			tbFile.Text = p1.fileNumber;
			tbSurplus.Text = p1.surplusAmount.ToString();
			tbState.Text = p1.state;
			tbCounty.Text = p1.county;

			dateReviewed.SelectedDate = p1.dateReviewed;
			tbResearcher.Text = p1.researcher;
			tbHeld.Text = p1.verifyBy;
			tbVerify.Text = p1.verifyHow;
			tbAddress.Text = p1.foreclosedAdd;
			dateRecorded.SelectedDate = p1.dateRecorded;
			tbBook.Text = p1.book;
			tbPage.Text = p1.page;
			dateForclosed.SelectedDate = p1.dateForclosed;
			tbNote.Text = p1.amountOnNote.ToString();
			tbBeneficiary.Text = p1.beneficiary;


			if (p1.hasEstatesForm)
				cbAddEstates.IsChecked = true;

			tbEstatesNameDeceased.Text = p1.mtgName.Split(new char[] { ',' })[0];
			tbEstatesFileNumber.Text = p1.fileNumber;
			tbEstatesFunds.Text = p1.surplusAmount.ToString();
			cbEstatesCounty.Text = p1.county;
			cbEstatesState.Text = p1.state;
			dtEstateDateReviewed.SelectedDate = p1.dateReviewed;
			tbEstatesResearcher.Text = p1.researcher;
			tbEstatePartOf.Text = p1.estatePartOf;
			tbEstatesHow.Text = p1.verifyHow + " " + p1.verifyBy;



			tbEstatesDeathAtt.Text = p1.pdfEstateDeathCert;
			tbEstatesWillAtt.Text = p1.pdfEstateWill;
			tbEstatesNotifAtt.Text = p1.pdfEstateNotif;

			foreach (string itm in p1.pdfEstateAnyDocsList)
			{
				listEstateAnyDocs.Items.Add(itm);
			}

			foreach (string itm in p1.pdfEstateAnyOtherDocsList)
			{
				listEstateAnyOtherDocs.Items.Add(itm);
			}
			try
			{
				tbEstatesName1.Text = p1.conclusionList[0].Name;
				tbEstatesAmount1.Text = p1.conclusionList[0].Amount;

				tbEstatesName2.Text = p1.conclusionList[1].Name;
				tbEstatesAmount2.Text = p1.conclusionList[1].Amount;

				tbEstatesName3.Text = p1.conclusionList[2].Name;
				tbEstatesAmount3.Text = p1.conclusionList[2].Amount;

				tbEstatesName4.Text = p1.conclusionList[3].Name;
				tbEstatesAmount4.Text = p1.conclusionList[3].Amount;
			}
			catch{}

			mtgList.Clear();
			judList.Clear();
			claimantList.Clear();
			conclusionList.Clear();
			
			List<mtgClass> mtgList2 = new List<mtgClass>();
			List<judgementClass> judList2 = new List<judgementClass>();
			List<conclusionClass> conclusionList2 = new List<conclusionClass>();
			List<claimantsClass> claimantList2 = new List<claimantsClass>();

			mtgList2 = p1.mtgList;
			judList2 = p1.judgmentList;
			claimantList2 = p1.claimantList;
			conclusionList2 = p1.conclusionList;
			
			foreach (mtgClass mtg in mtgList2)
			{
				AddMtgs(mtg);
			}

			foreach (claimantsClass cls in claimantList2)
			{
				AddClaimants(cls);
			}

			foreach (judgementClass jud in judList2)
			{
				AddJudgCases(jud);
			}

			foreach (conclusionClass conCls in conclusionList2)
			{
				AddConclusion(conCls);
			}
			fs.Close();
		}


		private void Settings_Click(object sender, RoutedEventArgs e)
		{
			Settings set = new Settings();
			set.ShowDialog();
		}

		private void button1_Click_2(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";
			if (opendialog.ShowDialog() == true)
			{
				mtgAttButton.Content = opendialog.FileName;
			}
		}

		private void button2_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";
			if (opendialog.ShowDialog() == true)
			{
				judAttButton.Content = opendialog.FileName;
			}
		}

		private void button11_Click_2(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";
			if (opendialog.ShowDialog() == true)
			{
				conclusionAttButton.Content = opendialog.FileName;
			}
		}

		public static void CombineMultiplePDFs(string[] fileNames, string outFile)
		{
			try
			{
				iTextSharp.text.Document document = new iTextSharp.text.Document();
				PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
				if (writer == null)
				{
					return;
				}
				document.Open();

				foreach (string fileName in fileNames)
				{
					if (System.IO.File.Exists(fileName))
					{
						PdfReader reader = new PdfReader(fileName);
						reader.ConsolidateNamedDestinations();
						for (int i = 1; i <= reader.NumberOfPages; i++)
						{
							PdfImportedPage page = writer.GetImportedPage(reader, i);
							writer.AddPage(page);
						}

						PRAcroForm form = reader.AcroForm;
						if (form != null)
						{
							writer.CopyDocumentFields(reader);
						}
						reader.Close();
					}
				}
				writer.Close();
				document.Close();

			}
			catch
			{
				MessageBox.Show("Close the pdf file and try again.");
			}

		}

		private void button30_Click(object sender, RoutedEventArgs e)
		{
			openDialogPDF(textBox30, pdfCheck1);
		}

		private void button31_Click(object sender, RoutedEventArgs e)
		{
			openDialogPDF(textBox31, pdfCheck2);
		}

		private void button32_Click(object sender, RoutedEventArgs e)
		{
			openDialogPDF(textBox32, pdfCheck3);
		}

		private void button33_Click(object sender, RoutedEventArgs e)
		{
			openDialogPDF(textBox33, pdfCheck5);
		}

		private void button34_Click(object sender, RoutedEventArgs e)
		{
			openDialogPDF(textBox34, pdfCheck6);
		}

		private void button35_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";
			if (opendialog.ShowDialog() == true)
			{
				if (!listBoxPDF.Items.Contains(opendialog.FileName))
					listBoxPDF.Items.Add(opendialog.FileName);
				pdfMisc = true;
			}
		}


		private void openDialogPDF(TextBox tb, bool markBool)
		{
			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";


			if (opendialog.ShowDialog() == true)
			{
				tb.Text = opendialog.FileName;
				markBool = true;
			}
		}

		private void XmlSave()
		{


			p1.miscPdfs = tempMisc;

			p1.mtgName = tbName.Text;
			p1.deedHolder = tbDeedHolder.Text;
			p1.fileNumber = tbFile.Text;

			if (tbSurplus.Text.Length > 0)
				p1.surplusAmount = Convert.ToDouble(tbSurplus.Text);
			else
				p1.surplusAmount = 0;

			p1.county = tbCounty.Text;
			p1.state = tbState.Text;
			p1.dateReviewed = dateReviewed.SelectedDate.Value;
			p1.researcher = tbResearcher.Text;
			p1.verifyHow = tbVerify.Text;
			p1.verifyBy = tbHeld.Text;
			p1.foreclosedAdd = tbAddress.Text;
			p1.dateRecorded = dateRecorded.SelectedDate.Value;
			p1.book = tbBook.Text;
			p1.page = tbPage.Text;
			p1.dateForclosed = dateForclosed.SelectedDate.Value;

			if (tbNote.Text.Length > 0)
				p1.amountOnNote = Convert.ToDouble(tbNote.Text);
			else
				p1.amountOnNote = 0;

			p1.beneficiary = tbBeneficiary.Text;




			p1.estatePartOf = tbEstatePartOf.Text;

			p1.pdfEstateAnyDocsList = new List<string>();
			p1.pdfEstateAnyOtherDocsList = new List<string>();

			foreach (string itm in anyDocsList)
			{
				p1.pdfEstateAnyDocsList.Add(itm);
			}

			foreach (string itm in anyOtherDocsList)
			{
				p1.pdfEstateAnyOtherDocsList.Add(itm);
			}


			p1.hasEstatesForm = cbAddEstates.IsChecked.Value;


			p1.mtgList = mtgList;
			p1.judgmentList = judList;
			p1.conclusionList = conclusionList;
			p1.claimantList = claimantList;

			string xmlPath;
			if (isMtgForm)
				xmlPath = documentDir + "files\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + ".p1";
			if (isDeedsForm)
				xmlPath = documentDir + "files\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + ".p3";
			else
				xmlPath = documentDir + "files\\" + tbName.Text.Substring(0, tbName.Text.IndexOf(',')).ToUpper() + ".p2";
			XmlSerializer serializer = new XmlSerializer(p1.GetType());
			if (File.Exists(xmlPath))
				File.Delete(xmlPath);
			FileStream fs = new FileStream(xmlPath, FileMode.Create);
			serializer.Serialize(fs, p1);

			fs.Close();
		}

		private void tbState_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (tbState.SelectedItem != null)
			{
				tbCounty.Items.Clear();
				foreach (Tuple<string, string> item in CountyStateList.FindAll(x => x.Item2 == tbState.SelectedItem.ToString()))
				{
					tbCounty.Items.Add(item.Item1);
				}
			}
		}


		public class project1
		{
			public string pdfForclosureOrder { get; set; }
			public string pdfDisbursement { get; set; }
			public string pdfNotification { get; set; }
			public string pdfDeed { get; set; }

			public string pdfSheriffDeed { get; set; }

			public List<string> miscPdfs { get; set; }
			public string mtgName { get; set; }
			public string deedHolder { get; set; }
			public string fileNumber { get; set; }
			public double surplusAmount { get; set; }
			public string county { get; set; }
			public string state { get; set; }
			public DateTime dateReviewed { get; set; }
			public string researcher { get; set; }
			public string verifyHow { get; set; }
			public string verifyBy { get; set; }
			public string foreclosedAdd { get; set; }
			public DateTime dateRecorded { get; set; }
			public DateTime dateForclosed { get; set; }
			public string book { get; set; }
			public string page { get; set; }
			public double amountOnNote { get; set; }
			public string beneficiary { get; set; }

			public List<mtgClass> mtgList { get; set; }
			public List<judgementClass> judgmentList { get; set; }
			public List<conclusionClass> conclusionList { get; set; }
			public List<claimantsClass> claimantList { get; set; }
			public bool hasEstatesForm { get; set; }

			public string estatePartOf { get; set; }

			public string pdfEstateDeathCert { get; set; }
			public string pdfEstateWill { get; set; }
			public string pdfEstateNotif { get; set; }

			public List<string> pdfEstateAnyDocsList { get; set; }

			public List<string> pdfEstateAnyOtherDocsList { get; set; }

		}

		private void mtgDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (mtgDataGrid.SelectedIndex > -1)
			{

				isMtgLoaded = true;

				ClearEntryMtg();

				mtgClass tempMtg = (mtgClass)mtgDataGrid.SelectedItem;

				selectedID = tempMtg.ID;

				comboBox1.Text = tempMtg.Satisfied;



				mtgName.Text = tempMtg.Name;
				mtgAmount.Text = tempMtg.Amount;
				mtgLoanNum.Text = tempMtg.LoanNum;
				mtgDate.SelectedDate = tempMtg.Date;

				mtgBook.Text = tempMtg.Book;
				mtgPage.Text = tempMtg.Page;
				mtgAddress.Text = tempMtg.Address;

				mtgBookSatisfied.Text = tempMtg.BookSatisfied;
				mtgPageSatisfied.Text = tempMtg.PageSatisfied;

				mtgAttButton.Content = tempMtg.Attachment;

				if (mtgAttButton.Content.ToString().Length < 3)
				{
					mtgAttButton.Content = "ATTACHMENT (PDF)";
				}


			}

		}

		private void judDataGrid_DoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (judDataGrid.SelectedIndex > -1)
			{
				isJudLoaded = true;

				ClearEntryJud();

				judgementClass judTemp = (judgementClass)judDataGrid.SelectedItem;

				selectedID = judTemp.ID;

				judCaseNum.Text = judTemp.CaseNum;
				judDateRec.SelectedDate = judTemp.DateRec;
				judOrigAmount.Text = judTemp.OriginalAmount;
				judCurrAmount.Text = judTemp.CurrentAmount;
				judContact.Text = judTemp.Contact2;
				judName.Text = judTemp.Name;

				judAttButton.Content = judTemp.Attachment;

				if (judAttButton.Content.ToString().Length < 3)
					judAttButton.Content = "ATTACHMENT (PDF)";

			}

		}
		private void conclusionDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (conclusionDataGrid.SelectedIndex > -1)
			{
				isConcLoaded = true;

				ClearEntryConclusion();

				conclusionClass conCls = (conclusionClass)conclusionDataGrid.SelectedItem;

				selectedID = conCls.ID;
				tbConclusionName.Text = conCls.Name;
				tbConclusionAmount.Text = conCls.Amount;
				conclusionDate.SelectedDate = conCls.Date;

				conclusionAttButton.Content = conCls.Attachment;

				if (conclusionAttButton.Content.ToString().Length < 3)
					conclusionAttButton.Content = "ATTACHMENT (PDF)";
			}
		}

		private void claimantDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (claimantsDataGrid.SelectedIndex > -1)
			{
				isClaimantLoaded = true;

				ClearEntryClaimant();

				claimantsClass clCls = (claimantsClass)claimantsDataGrid.SelectedItem;

				selectedID = clCls.ID;
				tbClaimantName.Text = clCls.ClaimantName;
				tbClaimantPhone.Text = clCls.ClaimantPhone;
				tbClaimantAddress.Text = clCls.ClaimantAddress;
				tbClaimantGuess.Text = clCls.ClaimantGuess;
			}
		}

		private void button36_Click(object sender, RoutedEventArgs e)
		{
			if (listBoxPDF.SelectedIndex > -1)
			{
				listBoxPDF.Items.RemoveAt(listBoxPDF.SelectedIndex);
			}
		}

		private void CheckBox_Checked(object sender, RoutedEventArgs e)
		{
			expander8.Visibility = System.Windows.Visibility.Visible;
		}

		private void estatesUnchecked(object sender, RoutedEventArgs e)
		{
			expander8.Visibility = System.Windows.Visibility.Collapsed;
		}
		private void estateNext_Click(object sender, RoutedEventArgs e)
		{
			expander8.IsExpanded = false;
			expander6.IsExpanded = true;
		}

		private void estatesState_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (cbEstatesState.SelectedItem != null)
			{
				cbEstatesCounty.Items.Clear();
				foreach (Tuple<string, string> item in CountyStateList.FindAll(x => x.Item2 == cbEstatesState.SelectedItem.ToString()))
				{
					cbEstatesCounty.Items.Add(item.Item1);
				}
			}
		}

		private void attCLR_Click(object sender, RoutedEventArgs e)
		{
			mtgAttButton.Content = "ATTACHMENT (PDF)";
		}

		private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			TextBox tb = (TextBox)sender;
			if (tb.Text.Length > 0)
			{
				double res;
				if (double.TryParse(tb.Text, out res) == false)
				{
					MessageBox.Show("The value is not in the correct format!");
				}
			}
		}

		private void tbFile_TextChanged(object sender, TextChangedEventArgs e)
		{
			tbEstatesFileNumber.Text = tbFile.Text;
		}

		private string GetPdfAtt()
		{
			string attPath = "";

			OpenFileDialog opendialog = new OpenFileDialog();
			opendialog.Filter = "PDF Files|*.pdf";
			opendialog.Title = "PDF ATTACHMENT";


			if (opendialog.ShowDialog() == true)
			{
				attPath = opendialog.FileName;
			}

			return attPath;
		}

		private void deathCertButton_Click(object sender, RoutedEventArgs e)
		{
			tbEstatesDeathAtt.Text = GetPdfAtt();
		}
		
		private void willButton_Click(object sender, RoutedEventArgs e)
		{
			tbEstatesWillAtt.Text = GetPdfAtt();
		}

		private void anyDocsButton_Click(object sender, RoutedEventArgs e)
		{
			string item = GetPdfAtt();

			if (!listEstateAnyDocs.Items.Contains(item))
				listEstateAnyDocs.Items.Add(item);
		}

		private void notifButton_Click(object sender, RoutedEventArgs e)
		{
			tbEstatesNotifAtt.Text = GetPdfAtt();
		}

		private void anyOtherDocsButton_Click(object sender, RoutedEventArgs e)
		{
			string item = GetPdfAtt();

			if (!listEstateAnyOtherDocs.Items.Contains(item))
				listEstateAnyOtherDocs.Items.Add(item);
		}

		private void delButton1_Click(object sender, RoutedEventArgs e)
		{
			if (listEstateAnyDocs.SelectedIndex >= 0)
			{
				listEstateAnyDocs.Items.RemoveAt(listEstateAnyDocs.SelectedIndex);
			}
		}

		private void delButton2_Click(object sender, RoutedEventArgs e)
		{
			if (listEstateAnyOtherDocs.SelectedIndex >= 0)
			{
				listEstateAnyOtherDocs.Items.RemoveAt(listEstateAnyOtherDocs.SelectedIndex);
			}
		}

		private void delAttachment(object sender, RoutedEventArgs e)
		{
			Button tempBut = (Button)sender;

			if (tempBut == delButtonAtt1)
				textBox30.Text = "";
			if (tempBut == delButtonAtt2)
				textBox31.Text = "";
			if (tempBut == delButtonAtt3)
				textBox32.Text = "";
			if (tempBut == delButtonAtt4)
				textBox33.Text = "";
			if (tempBut == delButtonAtt5)
				textBox34.Text = "";

			if (tempBut == delButtonAtt6)
				tbEstatesDeathAtt.Text = "";
			if (tempBut == delButtonAtt7)
				tbEstatesWillAtt.Text = "";
			if (tempBut == delButtonAtt8)
				tbEstatesNotifAtt.Text = "";

		}

		private void dateReviewed_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
		{
			dtEstateDateReviewed.SelectedDate = dateReviewed.SelectedDate;
		}


		private void deeds_Click(object sender, RoutedEventArgs e)
		{
			new DeedsWindow().Show();
			Close();
		}
	}
}
