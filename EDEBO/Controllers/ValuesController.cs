using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using EDEBO.Repository;
using System.ComponentModel;
using Newtonsoft.Json.Linq;
using System.Dynamic;
using System.Web;
using MongoDB.Bson;
using System.IO;
using System.Text;
using MongoDB.Bson.Serialization;
using EDEBO.Services;
using Newtonsoft.Json;

namespace EDEBO.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        private readonly IStudentRepository _studentRepository;
        private readonly IProcessService _processService;

        public ValuesController(IStudentRepository studentRepository,
                                IProcessService processService)
        {
            _studentRepository = studentRepository;
            _processService = processService;
        }

        [HttpGet]
        [Route("get-users")]
        public async Task<IEnumerable<object>> GetUsers()
        {
            return await _studentRepository.GetAllStudents();
        }

        [HttpPost]
        [Route("add-student")]
        public async Task<IActionResult> AddStudent()
        {
            using (StreamReader reader = new StreamReader(Request.Body, Encoding.UTF8))
            {
                var json = reader.ReadToEnd();
                // convert to dictionary
                var dict = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(json);
                // add hash
                dict = _processService.addHashToDictionary(dict);
                // push to DB
                await _studentRepository.AddStudent(dict.ToBsonDocument());
                return Ok();
            }
        }

        [HttpPost]
        [Route("add-students")]
        public async Task<IActionResult> AddStudents()
        {
            using (StreamReader reader = new StreamReader(Request.Body, Encoding.UTF8))
            {
                var json = reader.ReadToEnd();
                // convert to dictionary
                var dictList = JsonConvert.DeserializeObject<List<Dictionary<string, dynamic>>>(json);
                // add hash
                dictList = _processService.addHashToDictionaryList(dictList);
                // serialize to JSON
                json = JsonConvert.SerializeObject(dictList);
                // get BSON Document
                var document = BsonSerializer.Deserialize<List<BsonDocument>>(json);
                // push to DB
                await _studentRepository.AddStudents(document);
                return Ok();
            }
        }

        [HttpPost]
        [Route("update-student/{studentId}")]
        public async Task<IActionResult> UpdateStudent(string studentId)
        {
            try
            {
                using (StreamReader reader = new StreamReader(Request.Body, Encoding.UTF8))
                {
                    var findObj = (IDictionary<string, dynamic>) await _studentRepository.GetStudent(studentId);
                    var oldStudent = new Dictionary<string, dynamic>(findObj);
                    var json = reader.ReadToEnd();
                    var newStudent = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(json);
                    newStudent = _processService.updateValue(oldStudent, newStudent);

                    await _studentRepository.UpdateStudent(studentId, newStudent.ToBsonDocument());
                    return Ok();
                }
            } 
            catch (Exception e)
            {
                return BadRequest();
            }
        }


        [HttpGet]
        [Route("test-students")]
        public async Task<IActionResult> TestStudents()
        {
            try
            {
                    var json = getObjects(100000).ToJson();
                    var documents = BsonSerializer.Deserialize<List<BsonDocument>>(json);

                    string[] studentIds = documents.Select(i => i.GetElement(StudentRepository.uniqueField).Value.ToString()).ToArray();
                    var oldStudents = await _studentRepository.GetStudentsByIds(studentIds);
                    var oldStudentsListDict = new List<Dictionary<string, object>>();
                    foreach (var student in oldStudents)
                    {
                        oldStudentsListDict.Add(new Dictionary<string, dynamic>((IDictionary<string, dynamic>)student));
                    }


                    var newStudentsListDict = JsonConvert.DeserializeObject<List<Dictionary<string, dynamic>>>(json);

                    var updateList = _processService.updateValueList(oldStudentsListDict, newStudentsListDict);

                    foreach (var update in updateList)
                    {
                        _studentRepository.UpdateStudent(update[StudentRepository.uniqueField], update.ToBsonDocument());
                    }


                    return Ok();
            }
            catch (Exception e)
            {
                return BadRequest();
            }
        }

        [HttpPost]
        [Route("update-students")]
        public async Task<IActionResult> UpdateStudents(string studentId)
        {
            try
            {
                using (StreamReader reader = new StreamReader(Request.Body, Encoding.UTF8))
                {
                    var json = await reader.ReadToEndAsync();
                    var documents = BsonSerializer.Deserialize<List<BsonDocument>>(json);

                    string[] studentIds = documents.Select(i => i.GetElement(StudentRepository.uniqueField).Value.ToString()).ToArray();
                    var oldStudents = await _studentRepository.GetStudentsByIds(studentIds);
                    var oldStudentsListDict = new List<Dictionary<string, object>>();
                    foreach (var student in oldStudents)
                    {
                        oldStudentsListDict.Add(new Dictionary<string, dynamic>((IDictionary<string, dynamic>)student));
                    }

                   
                    var newStudentsListDict = JsonConvert.DeserializeObject<List<Dictionary<string, dynamic>>>(json);

                    var updateList = _processService.updateValueList(oldStudentsListDict, newStudentsListDict);

                    foreach(var update in updateList)
                    {
                        _studentRepository.UpdateStudent(update[StudentRepository.uniqueField], update.ToBsonDocument());
                    }


                    return Ok();
                }
            }
            catch (Exception e)
            {
                return BadRequest();
            }
        }

        public static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
                expandoDict[propertyName] = propertyValue;
            else
                expandoDict.Add(propertyName, propertyValue);
        }

        private List<dynamic> getObjects(int count)
        {
            var res = new List<dynamic>();
            for(int i = 0; i < count; i++)
            {
                res.Add(new {
                    studentId = "studentId" + i,
                    field1 = "field1val-" + i,
                    field2 = "field2val-" + i,
                    field3 = "field3val-" + i,
                    field4 = "field4val-" + i,
                    field5 = "field5val-" + i,
                    field6 = "field6val-" + i,
                    field7 = "field7val-" + i,
                    field8 = "field8val-" + i,
                    field9 = "field9val-" + i,
                    field10 = "field10val-" + i,
                    field11 = "field11val" + i,
                    field12 = "field12val-" + i,
                    field13 = "field13val" + i,
                    field14 = "field14val" + i,
                    field15 = "field15val" + i,
                    field16 = "field16val" + i,
                    field17 = "field17val" + i,
                    field18 = "field18val" + i,
                    field19 = "field19val-" + i,
                    field20 = "field20val" + i,
                    field21 = "field21val" + i,
                    field22 = "field22val" + i,
                    field23 = "field23val" + i,
                    field24 = "field24val-" + i,
                    field25 = "field25val" + i,
                    field26 = "field26val" + i,
                    field27 = "field27val-" + i,
                    field28 = "field28val" + i,
                    field29 = "field29val" + i,
                    field30 = "field30val" + i,
                    field31 = "field31val" + i,
                    field32 = "field32val" + i,
                    field33 = "field33val-" + i,
                    field34 = "field34val" + i,
                    field35 = "field35val" + i,
                    field36 = "field36val" + i,
                    field37 = "field37val" + i,
                    field38 = "field38val-" + i,
                    field39 = "field39val" + i
                });
            }
            return res;
        }


        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public ActionResult<string> Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
