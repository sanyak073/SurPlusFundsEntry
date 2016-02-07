using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace SurplusFundsEntry
{
	/// <summary>
	/// Interaction logic for DeedsWindow.xaml
	/// </summary>
	public partial class DeedsWindow : Window
	{
		public static readonly string Browse = "browse...";
		private static List<Tuple<string, string>> _countyStateList = new List<Tuple<string, string>>();

		public DeedsWindow()
		{
			InitializeComponent();

			DeedRecognitionDatePicker.SelectedDate = DateTime.Now;
			DateReviewed.SelectedDate = DateTime.Now;
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWb = xlApp.Workbooks.Open(Properties.Settings.Default.pathTemplate, false, false);
			Excel.Worksheet xlWs = xlWb.Worksheets["countylist"];

			StateComboBox.Items.Add("");
			for (int rowx = 2; rowx <= xlWs.UsedRange.Rows.Count; rowx++)
			{
				if (!StateComboBox.Items.Contains(xlWs.Cells[rowx, 2].Value2))
				{
					StateComboBox.Items.Add(xlWs.Cells[rowx, 2].Value2);
					AddressStateComboBox.Items.Add(xlWs.Cells[rowx, 2].Value2);
				}
				_countyStateList.Add(new Tuple<string, string>(xlWs.Cells[rowx, 1].value, xlWs.Cells[rowx, 2].Value2));
			}

			xlWb.Close(false);
			xlApp.Quit();
		}
		/* Attachments methods
		private void ForeClosureOrderButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void ForeClosureOrderRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void DisbursementSheetButtonBrowse_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void DisbursementSheetRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void NotificationAddressBrowseButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void NotificationAddressRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void DeedBrowseButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void DeedRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void SheriffDeedBrowseButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void SheriffDeedRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void MiscBrowseButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void MiscRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}
		*/

		private void NamesOfPersonsTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			if (NamesOfPersonsTextBox.Text.Length > 0)
			{
				char[] toReplace = { '\\', '/', ':', '*', '?', '\"', '<', '>', '|' };
				foreach (char chr in toReplace)
					NamesOfPersonsTextBox.Text = NamesOfPersonsTextBox.Text.Replace(chr.ToString(), "");
			}
		}

		private void MortgageHistoryExpander_Loaded(object sender, RoutedEventArgs e)
		{
			MortgageHolderNameTextBox.Focus();
		}

		private void MtgBookSatisfied_TextChanged(object sender, TextChangedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void mtgBookPage_TextChanged(object sender, TextChangedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void Button5_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void button12_Click(object sender, RoutedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void dataGrid1_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void mtgDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			throw new System.NotImplementedException();
		}

		private void SaveAddButton_Click(object sender, RoutedEventArgs e)
		{
			//AddMtgs(selectedID, mtgName.Text, mtgAmount.Text, mtgLoanNum.Text, mtgDate.SelectedDate.Value, mtgBook.Text, mtgPage.Text, mtgAddress.Text, satisfied, pdf, mtgBookSatisfied.Text, mtgPageSatisfied.Text, comboBox1.Text);
		}

		private void StateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (StateComboBox.SelectedItem != null)
			{
				StateComboBox.Items.Clear();
				foreach (Tuple<string, string> item in _countyStateList.FindAll(x => x.Item2 == StateComboBox.SelectedItem.ToString()))
				{
					StateComboBox.Items.Add(item.Item1);
				}
			}
		}

		private void AnotherBank_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			ComboBoxItem cbItem = (ComboBoxItem)AnotherBankComboBox.SelectedItem;
			if (cbItem != null)
			{
				if (cbItem.Content.ToString() != "YES")
				{
					WhoHasLoanNumberLabel.Visibility = System.Windows.Visibility.Visible;
					WhoHasLoanNumberTextBox.Visibility = System.Windows.Visibility.Visible;
				}
				else
				{
					WhoHasLoanNumberLabel.Visibility = System.Windows.Visibility.Hidden;
					WhoHasLoanNumberTextBox.Visibility = System.Windows.Visibility.Hidden;
				}
			}
		}
	}
}
