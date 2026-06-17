using Microsoft.Data.SqlClient;
using Retail_Crawler.Models;

namespace Retail_Crawler.Services
{
    public class ReportDAO
    {
        private readonly string connectionString;

        public ReportDAO(IConfiguration configuration)
        {
            connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("DefaultConnection is missing.");
        }

        internal List<StoreModel> GetAllStores()
        {
            List<StoreModel> storeList = new List<StoreModel>();
            string commandString = "SELECT * FROM dbo.Stores";
            using (SqlConnection connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand(commandString, connection))
            {
                try
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            storeList.Add(
                                new StoreModel
                                {
                                    Id = reader.GetInt32(0),
                                    StoreNumber = reader.GetString(1)
                                });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    throw;
                }
            }
            return storeList;
        }
    }
}
