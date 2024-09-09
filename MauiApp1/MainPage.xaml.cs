using Microsoft.Maui.Controls;
using System;
using System.Collections.ObjectModel;

namespace MauiApp1
{
    public partial class MainPage : ContentPage
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public ObservableCollection<SearchResult> SearchResults { get; set; }

        public MainPage()
        {
            InitializeComponent();

            // Initialize dates
            StartDate = DateTime.Today;
            EndDate = DateTime.Today;



            // Initialize SearchResults collection
            SearchResults = new ObservableCollection<SearchResult>();

			            // Set DataContext for data binding
            BindingContext = this;
        }

        private void OnSearchClicked(object sender, EventArgs e)
        {
            // Example search logic - replace with your own
            string searchText = txtSearch.Text;
            DateTime startDate = dpStartDate.Date;
            DateTime endDate = dpEndDate.Date;

            if (endDate < startDate)
            {
                DisplayAlert("Error", "End date cannot be before start date.", "OK");
                return;
            }

            // Simulate search results
            SearchResults.Clear();
            for (int i = 0; i < 10; i++)
            {
                SearchResults.Add(new SearchResult
                {
                    Id = (i + 1).ToString(),
                    Name = $"Item {i + 1}",
                    Date = DateTime.Now.AddDays(i).ToShortDateString()
                });
            }
        }
    }

    public class SearchResult
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? Date { get; set; }
    }
}
