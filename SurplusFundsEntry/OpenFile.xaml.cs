using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;

namespace SurplusFundsEntry
{
	/// <summary>
	/// Interaction logic for OpenFile.xaml
	/// </summary>
	public partial class OpenFile : Window
	{
		public OpenFile()
		{
			InitializeComponent();
		}

		public static string FileToOpen;
		public static List<pDoc> docsList = new List<pDoc>();
		public static List<pDoc> findList = new List<pDoc>();

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			docsList.Clear();
			findList.Clear();

			foreach (string filex in Directory.GetFiles(Properties.Settings.Default.pathSaveFolder, "*", SearchOption.AllDirectories))
			{
				string file = filex.ToUpper();

				if (file.EndsWith(".P1") && !file.Contains("~$"))
				{
					List<string> splitList = file.Split(new char[] { '\\' }).ToList();
					string fullpath = file;
					string name = splitList[splitList.Count - 1].Replace(".P1", "");
					string state = splitList[splitList.Count - 6];
					string county = splitList[splitList.Count - 5];
					string type = "MORTGAGE";
					
					docsList.Add(new pDoc() { fullpath = fullpath, Name = name, County = county, State = state, Type = type });
				}
				searchTextBox.Focus();
			}

			foreach (string filex in Directory.GetFiles(Properties.Settings.Default.pathSaveFolder, "*", SearchOption.AllDirectories))
			{
				string file = filex.ToUpper();
				if (file.EndsWith(".P2") && !file.Contains("~$"))
				{
					List<string> splitList = file.Split(new char[] { '\\' }).ToList();
					string fullpath = file;
					string name = splitList[splitList.Count - 1].Replace(".P2", "");
					string state = splitList[splitList.Count - 6];
					string county = splitList[splitList.Count - 5];
					string type = "TAX";
					docsList.Add(new pDoc() { fullpath = fullpath, Name = name, County = county, State = state, Type = type });
				}
			}

			foreach (pDoc doc in docsList)
			{
				dataGrid1.Items.Add(doc);
			}
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			dataGrid1.Items.Clear();
			findList.Clear();
			foreach (pDoc doc in docsList)
			{
				dataGrid1.Items.Add(doc);
			}

		}

		private void dataGrid1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			pDoc doc = (pDoc)dataGrid1.SelectedItem;
			MainWindow.FileToOpen = doc.fullpath;

			if (doc.Type == "MORTGAGE")
			{
				MainWindow.isMtgForm = true;
				MainWindow.isTaxForm = false;
			}
			else
			{
				MainWindow.isMtgForm = false;
				MainWindow.isTaxForm = true;
			}

			this.Close();
		}

		private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			findList.Clear();
			dataGrid1.Items.Clear();

			findList.AddRange(docsList.FindAll(x => x.County.Contains(searchTextBox.Text.ToUpper())));
			findList.AddRange(docsList.FindAll(x => x.Name.Contains(searchTextBox.Text.ToUpper())));
			findList.AddRange(docsList.FindAll(x => x.State.Contains(searchTextBox.Text.ToUpper())));
			findList.AddRange(docsList.FindAll(x => x.Type.Contains(searchTextBox.Text.ToUpper())));

			findList = findList.Distinct().ToList();
			foreach (pDoc doc in findList)
			{
				dataGrid1.Items.Add(doc);
			}

		}

		private void delButton_Click(object sender, RoutedEventArgs e)
		{
			if (dataGrid1.SelectedIndex > -1)
			{
				pDoc delDoc = (pDoc)dataGrid1.Items[dataGrid1.SelectedIndex];
				MessageBoxResult mbres = MessageBox.Show("Are you sure you want to delete '" + delDoc.Name + "'?", "Delete", MessageBoxButton.YesNo);

				if (mbres == MessageBoxResult.Yes)
				{
					try
					{
						Directory.Delete(delDoc.fullpath.Substring(0, delDoc.fullpath.IndexOf("FORMS") + 5) + "\\" + delDoc.Name, true);
						dataGrid1.Items.RemoveAt(dataGrid1.SelectedIndex);
					}
					catch
					{
					}
				}
			}
		}
	}
	
	public class pDoc
	{
		public string fullpath { get; set; }
		public string Name { get; set; }
		public string State { get; set; }
		public string County { get; set; }
		public string Type { get; set; } //mtg or tax
	}
}