using Microsoft.Extensions.Options;
using MongoDB.Driver;

namespace EDEBO.Models
{
    public class StudentContext
    {
        private readonly IMongoDatabase _database = null;
        private readonly string _collection;
        public StudentContext(IOptions<Settings> settings)
        {
            var client = new MongoClient(settings.Value.ConnectionString);
            _database = client.GetDatabase(settings.Value.Database);
            _collection = settings.Value.Collection;
        }

        public IMongoCollection<dynamic> Students
        {
            get
            {
                return _database.GetCollection<dynamic>(_collection);
            }
        }
    }
}
