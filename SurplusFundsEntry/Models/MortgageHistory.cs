using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SurplusFundsEntry.Models
{
	public class MortgageHistory
	{
		public int ID { get; set; }
		public string Name { get; set; }
		public string Amount { get; set; }
		public string Loan { get; set; }
		public DateTime Date { get; set; }
		public string ForclEntity { get; set; }
		public string Book { get; set; }
		public string Page { get; set; }
		public string InterestedRate { get; set; }
		public string EstBalance { get; set; }
	}
}
