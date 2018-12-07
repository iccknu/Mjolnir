using System.Collections.Generic;

namespace EDEBO.Services
{
    public interface IProcessService
    {
        Dictionary<string, dynamic> addHashToDictionary(Dictionary<string, dynamic> dictionary);
        List<Dictionary<string, dynamic>> addHashToDictionaryList(List<Dictionary<string, dynamic>> dictionaryList);
        Dictionary<string, dynamic> updateValue(Dictionary<string, dynamic> oldVal, Dictionary<string, dynamic> newVal);
        List<Dictionary<string, dynamic>> updateValueList(List<Dictionary<string, dynamic>> oldValList,
            List<Dictionary<string, dynamic>> newValList);
    }
}
