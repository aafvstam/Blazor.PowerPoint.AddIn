/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
namespace Blazor.PowerPoint.AddIn.Components.Pages
{
    public partial class Weather
    {
        private WeatherForecast[]? forecasts;

        public bool IsLoading
        {
            get
            {
                return forecasts is null;
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await GetWeatherData();
        }

        private async Task RefreshButton()
        {
            await GetWeatherData();
        }

        private async Task GetWeatherData()
        {
            forecasts = null;

            // Simulate asynchronous loading to demonstrate streaming rendering
            await Task.Delay(500);

            var startDate = DateOnly.FromDateTime(DateTime.Now);
            var summaries = new[] { "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching" };
            forecasts = Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = startDate.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = summaries[Random.Shared.Next(summaries.Length)]
            }).ToArray();
        }

        private class WeatherForecast
        {
            public DateOnly Date { get; set; }
            public int TemperatureC { get; set; }
            public string? Summary { get; set; }
            public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);
        }
    }
}