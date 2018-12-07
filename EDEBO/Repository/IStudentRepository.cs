using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EDEBO.Repository
{
    public interface IStudentRepository
    {
        Task<IEnumerable<dynamic>> GetAllStudents();
        Task<dynamic> GetStudent(string id);
        Task AddStudent(dynamic student);
        Task AddStudents(dynamic students);
        Task<dynamic> UpdateStudent(string id, BsonDocument student);
        Task<IEnumerable<dynamic>> GetStudentsByIds(string[] studentIds);
    }
}
