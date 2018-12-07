using Microsoft.Extensions.Options;
using EDEBO.Models;
using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MongoDB.Bson.Serialization;

namespace EDEBO.Repository
{
    public class StudentRepository : IStudentRepository
    {
        private readonly StudentContext _context = null;
        private readonly string _collection;
        public const string uniqueField = "studentId";

        public StudentRepository(IOptions<Settings> settings)
        {
            _context = new StudentContext(settings);
            _collection = settings.Value.Collection;
        }

        public async Task AddStudent(dynamic student)
        {
            try
            {
                var collection = _context.Students.Database.GetCollection<BsonDocument>(_collection);
                await collection.InsertOneAsync(student);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task AddStudents(dynamic students)
        {
            try
            {
                var collection = _context.Students.Database.GetCollection<BsonDocument>(_collection);
                await collection.InsertManyAsync(students);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task<IEnumerable<dynamic>> GetAllStudents()
        {
            try
            {
                FilterDefinition<dynamic> filter = new BsonDocument();
                var cursor = await _context.Students
                   .FindAsync<dynamic>(filter);

                return cursor.ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task<dynamic> GetStudent(string id)
        {
            var filter = new BsonDocument(uniqueField, id);
            var cursor = await _context.Students.FindAsync<dynamic>(filter);
            return cursor.ToList()[0];
        }

        public async Task<dynamic> UpdateStudent(string id, BsonDocument student)
        {
            var filter = new BsonDocument(uniqueField, id);
            var collection = _context.Students.Database.GetCollection<BsonDocument>(_collection);
            var res = await collection.ReplaceOneAsync(filter, student, new UpdateOptions { IsUpsert = true });
            return res;
        }

        public async Task<IEnumerable<dynamic>> GetStudentsByIds(string[] studentIds)
        {
            var filter = Builders<BsonDocument>.Filter.In(uniqueField, studentIds);
            var collection = _context.Students.Database.GetCollection<BsonDocument>(_collection);
            var cursor = await collection.FindAsync<dynamic>(filter);
            return cursor.ToList();
        } 
    }
}
