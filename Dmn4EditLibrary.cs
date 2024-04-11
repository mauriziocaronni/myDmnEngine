using System.Xml.Linq;

// OLD 

namespace EditDmnLibrary
{
    public class Dmn4Edit {
        public class DmnObject 
        {
            public string fileName { get; set; }
            public Dictionary<string, string> inputDefs { get; set; }
            public Dictionary<string, string> outputDefs { get; set; }

            public List<Dictionary<string, string>> ruleTable { get; set; }
            public string decision { get; set; }

        }


        public static DmnObject CreateDmnObject (string fileName)
        {
            var dmnObject = new DmnObject();
            dmnObject.fileName = fileName;
            dmnObject.inputDefs = [];
            dmnObject.outputDefs = [];
            dmnObject.ruleTable = [];
            dmnObject.decision = "";

            var xdoc = XDocument.Load(fileName);
            var ns = xdoc.Root.Name.Namespace;
            var decision = xdoc.Descendants(ns + "decision").FirstOrDefault();
            dmnObject.decision = decision.Attribute("name").Value;

            var inputDefs = xdoc.Descendants(ns + "inputData");
            foreach (var inputDef in inputDefs)
            {
                dmnObject.inputDefs.Add(inputDef.Attribute("name").Value, inputDef.Attribute("typeRef").Value);
            }

            var outputDefs = xdoc.Descendants(ns + "outputData");
            foreach (var outputDef in outputDefs)
            {
                dmnObject.outputDefs.Add(outputDef.Attribute("name").Value, outputDef.Attribute("typeRef").Value);
            }

            var rules = xdoc.Descendants(ns + "rule");
            foreach (var rule in rules)
            {
                var ruleRow = new Dictionary<string, string>();
                foreach (var inputDef in dmnObject.inputDefs)
                {
                    ruleRow.Add(inputDef.Key, rule.Descendants(ns + "inputEntry").Where(x => x.Attribute("id").Value == inputDef.Key).FirstOrDefault().Value);
                }
                foreach (var outputDef in dmnObject.outputDefs)
                {
                    ruleRow.Add(outputDef.Key, rule.Descendants(ns + "outputEntry").Where(x => x.Attribute("id").Value == outputDef.Key).FirstOrDefault().Value);
                }
                dmnObject.ruleTable.Add(ruleRow);
            }

            return dmnObject;
        }


    }
}  
// end 