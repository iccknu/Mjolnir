using EDEBO.Repository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EDEBO.Services
{
    public class ProcessService : IProcessService
    {
        private const char deleteFieldPrefix = '_';
        private const string hashKey = "hash";

        public Dictionary<string, dynamic> addHashToDictionary(Dictionary<string, dynamic> dictionary)
        {
            SortedDictionary<string, dynamic> sortedDictionary = new SortedDictionary<string, dynamic>(dictionary);
            StringBuilder commonString = new StringBuilder("");
            foreach (KeyValuePair<string, dynamic> entry in sortedDictionary)
            {
                commonString.Append(entry.Value);
            }

            dictionary.Add(hashKey, sha256Hash(commonString.ToString()));
            return dictionary;
        }

        public List<Dictionary<string, dynamic>> addHashToDictionaryList(List<Dictionary<string, dynamic>> dictionaryList)
        {
            dictionaryList.ForEach(i => addHashToDictionary(i) );
            return dictionaryList;
        }

        private static String sha256Hash(string value)
        {
            StringBuilder Sb = new StringBuilder();

            using (var hash = SHA256.Create())
            {
                Encoding enc = Encoding.UTF8;
                Byte[] result = hash.ComputeHash(enc.GetBytes(value));

                foreach (Byte b in result)
                    Sb.Append(b.ToString("x2"));
            }

            return Sb.ToString();
        }

        public Dictionary<string, dynamic> updateValue(Dictionary<string, dynamic> oldVal, Dictionary<string, dynamic> newVal)
        {
            newVal = addHashToDictionary(newVal);
            // if hash is equals - object is equals
            if ((oldVal[hashKey]).Equals(newVal[hashKey])) {
                return newVal;
            }
            var build = new Dictionary<string, dynamic>();

            // update existing fields
            foreach (KeyValuePair<string, dynamic> entry in oldVal)
            {
                if (newVal.ContainsKey(entry.Key))
                {
                    build.Add(entry.Key, newVal[entry.Key]);
                }
                else
                {
                    // add delete prefix for deleted prop
                    if (entry.Key[0] != deleteFieldPrefix)
                    {
                        build.Add(deleteFieldPrefix + entry.Key, entry.Value);
                    }
                    // store deleted property
                    else 
                    {
                        build.Add(entry.Key, entry.Value);
                    }
                }
            }
           // add new fields
           foreach(KeyValuePair<string, dynamic> entry in newVal)
            {
                if (!build.ContainsKey(entry.Key))
                {
                    build[entry.Key] = entry.Value;
                }
            }
           // remove immutable field
            build.Remove("_id");
            return build;
        }

        public List<Dictionary<string, dynamic>> updateValueList(List<Dictionary<string, dynamic>> oldValList,
            List<Dictionary<string, dynamic>> newValList)
        {
           var updateValues = new List<Dictionary<string, dynamic>>();
           for(int n = 0; n < newValList.Count(); n++)
            {
                var thePrevDoc = oldValList.FirstOrDefault(i => i[StudentRepository.uniqueField].Equals(newValList[n][StudentRepository.uniqueField]));
                if (thePrevDoc != null)
                {
                    updateValues.Add(updateValue(thePrevDoc, newValList[n]));
                } else
                {
                    newValList[n] = addHashToDictionary(newValList[n]);
                    updateValues.Add(newValList[n]);
                }
            }
           return updateValues;
        }
    }
}
