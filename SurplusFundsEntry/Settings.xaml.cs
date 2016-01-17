using System.Windows;
using Microsoft.Win32;
using Winforms = System.Windows.Forms;

namespace SurplusFundsEntry
{
	/// <summary>
	/// Interaction logic for OpenFile.xaml
	/// </summary>
	public partial class Settings : Window
	{
		public Settings()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			TemplateTextBox.Text = Properties.Settings.Default.pathTemplate;
			SaveFolderTextBox.Text = Properties.Settings.Default.pathSaveFolder;
		}

		private void TemplateButton_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog
			{
				Filter = "Excel files|*.xl*",
				Title = "EXCEL TEMPLATE"
			};

			if (opendialog.ShowDialog() == true)
			{
				Properties.Settings.Default.pathTemplate = opendialog.FileName;
				TemplateTextBox.Text = opendialog.FileName;
				Properties.Settings.Default.Save();
			}
		}

		private void DeedsButton_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog opendialog = new OpenFileDialog
			{
				Filter = "Excel files|*.xl*",
				Title = "EXCEL DEEDS"
			};

			if (opendialog.ShowDialog() == true)
			{
				Properties.Settings.Default.pathDeeds = opendialog.FileName;
				TemplateTextBox.Text = opendialog.FileName;
				Properties.Settings.Default.Save();
			}
		}
		private void SaveFolderButton_Click(object sender, RoutedEventArgs e)
		{
			using (var fdiag = new Winforms.FolderBrowserDialog())
			{
				if (fdiag.ShowDialog() == Winforms.DialogResult.OK)
				{
					Properties.Settings.Default.pathSaveFolder = fdiag.SelectedPath;
					SaveFolderTextBox.Text = fdiag.SelectedPath;
					Properties.Settings.Default.Save();
				}
			}
		}
	}
}
