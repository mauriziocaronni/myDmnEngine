using net.adamec.lib.common.dmn.engine.parser;
using net.adamec.lib.common.dmn.engine.engine.execution.context;
using net.adamec.lib.common.dmn.engine.engine.definition;
using net.adamec.lib.common.dmn.engine.engine.execution.result;
using net.adamec.lib.common.dmn.engine.utils;

using System.Xml.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Xml.Serialization;
using NPOI.HSSF.Record.PivotTable;
using System.Diagnostics;
using System.IO.Enumeration;

namespace MyDmnEngine
{
    public class MyDmn {


        public static void LogMessage(Boolean debug, string message)
        {
            var logfile = "logfile.txt";
            if (debug)
                Console.WriteLine(message);
            if (logfile != "")
            {
                using (StreamWriter sw = File.AppendText(logfile))
                {
                    sw.WriteLine(message);
                }
            }
        }

        public static void CheckDmnSyntax(Boolean debug, string fileName)
        {
            try
            {
                var ctx = MyDmn.CreateDmnCtx(fileName);
                var inputList = MyDmn.GetDmnInputs(debug,fileName);
                var decision = MyDmn.GetDmnDecision(debug,fileName);
                var outputList = MyDmn.GetDmnOutputs(debug, fileName);
            }
            catch (Exception e)
            {
                throw new DmnException("CheckDmnSyntax error: " + e.Message, e);
            }
        }


        // crea DMN context
        public static DmnExecutionContext CreateDmnCtx(string fileName) 
        {
            try
            {
                var model = DmnParser.Parse13ext(fileName);
                var def = DmnDefinitionFactory.CreateDmnDefinition(model);
                var ctx = DmnExecutionContextFactory.CreateExecutionContext(model);
                return ctx;
            }
            catch (Exception e)
            {
                throw new DmnException("CreateDmnCtx error: " + e.Message, e);
            }   

        }

       // eccezione nella lettura dei DMN
        public class DmnException : Exception 
        {   
            public DmnException(string message) : base(message) { }
            public DmnException(string message, Exception innerException) : base(message, innerException) { }
        }

        // restituisce la lista degli input del DMN
        public static List<string> GetDmnInputs(Boolean debug, string fileName)
        {
            var inputList = new List<string>();
            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                LogMessage(debug,"inputList = ");
 
                foreach (var input in doc.Descendants(ns + "input"))
                {
                    if (input.Attribute("label") != null)
                    {
                        var inputName = input.Attribute("label").Value;
                        if (!string.IsNullOrEmpty(inputName))
                        {
                            inputList.Add(inputName);
                        }
                    }
                }
                foreach (var input in inputList)
                    {
                       MyDmn.LogMessage(debug,"\t" + input);
                    }

                return inputList;
            }
            catch (Exception e)
            {
                throw new DmnException("GetDMNInputs error: " + e.Message, e);
            }
        }

        // restituisce la lista degli inputTypes  del DMN
        public static List<string> GetDmnInputTypes(Boolean debug, string fileName)
        {
            var inputTypesList = new List<string>();
            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                LogMessage(debug,"inputTypes = ");
 
                foreach (var input in doc.Descendants(ns + "inputExpression"))
                {
                    if (input.Attribute("typeRef") != null)
                    {
                        var inputName = input.Attribute("typeRef").Value;
                        if (!string.IsNullOrEmpty(inputName))
                        {
                            inputTypesList.Add(inputName);
                        }
                    }
                }
                foreach (var input in inputTypesList)
                    {
                       MyDmn.LogMessage(debug,"\t" + input);
                    }

                return inputTypesList;
            }
            catch (Exception e)
            {
                throw new DmnException("GetDMNInputTypes error: " + e.Message, e);
            }
        }



        // restituisce la lista degli output del DMN
        public static List<string> GetDmnOutputs(Boolean debug, string fileName)
        {
            var outputList = new List<string>();
            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                foreach (var output in doc.Descendants(ns + "output"))
                {
                    if (output.Attribute("label") != null)
                    {
                        var outputName = output.Attribute("label").Value;
                        if (!string.IsNullOrEmpty(outputName))
                        {
                            outputList.Add(outputName);
                        }
                    }
                }                   
                MyDmn.LogMessage(debug,"outputList = ");
                foreach (var output in outputList)
                    {
                       MyDmn.LogMessage(debug,"\t" + output);
                    }                 
                return outputList;
            }
            catch (Exception e)
            {
                throw new DmnException("GetDMNoutputs error: " + e.Message, e);
            }
        }

       // restituisce la lista di liste di regole
        public static List<List<string>> GetDmnRuleTable(Boolean debug, string fileName)
        {
            //var rule = new List<string>();
            var ruleTable = new List<List<string>>();

            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                foreach (var rule in doc.Descendants(ns + "rule"))
                {
                    var ruleList = new List<string>();
                    foreach (var input in rule.Descendants(ns + "inputEntry"))
                    {
                        var inputText = input.Descendants(ns + "text").FirstOrDefault();
                        ruleList.Add(inputText.Value);
                    }

                    foreach (var output in rule.Descendants(ns + "outputEntry"))
                    {
                        var outputText = output.Descendants(ns + "text").FirstOrDefault();
                        ruleList.Add(outputText.Value);
                    }
                    ruleTable.Add(ruleList);

                }                   
                MyDmn.LogMessage(debug,"Rules = ");
                foreach (var rule in ruleTable)
                    {
                       MyDmn.LogMessage(debug,"\t" + rule.ToString());
                    }                 
                return ruleTable;
            }
            catch (Exception e)
            {
                throw new DmnException("GetDMNoutputs error: " + e.Message, e);
            }
        }

    // restituisce il nome della decisione
        public static string GetDmnDecision(Boolean debug, string fileName)
        {
            string decision = "";
            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                var decisionElement = doc.Descendants(ns + "decision").FirstOrDefault();
                if (decisionElement.Attribute("name") != null)
                {
                    decision = decisionElement.Attribute("name").Value;
                    return decision;
                }
            }
            catch (Exception e)
            {
                throw new DmnException("GetDMNInputs error: " + e.Message, e);
            }
            LogMessage(debug,"decision = " + decision);
            return decision;
        }

        // crea test data generico
        public static Dictionary<string, string> CreateTestInput ( List<string> inputList)
        {
            var inputDictionary = new Dictionary<string, string>();
            foreach (var input in inputList)
            {
                inputDictionary[input] = "1";
            }
            return inputDictionary;
        }


        // crea test data user qualification
        public static Dictionary<string, string> CreateTestUserQualfications ()
        {
            var inputDictionary = new Dictionary<string, string>();

            inputDictionary["Task"] = "ED";
            inputDictionary["Azienda"] = "C&P";
            inputDictionary["Compagnia"] = "GEN";
            inputDictionary["DannoTipo"] = "Acqua Condotta";
            inputDictionary["ContraenteTipo"] = "AZI";
            inputDictionary["TipoDiIncarico"] = "PER";
            inputDictionary["Agenzia"] = "";
            inputDictionary["Agenzia"] = "";
         
            inputDictionary["Riserva"] = "20001";
            inputDictionary["Authority"] = "300000";
            inputDictionary["Broker"] = "";
            inputDictionary["Broker"] = "";
            inputDictionary["Amministratore"] = "";
            inputDictionary["Provincia"] = "MI";
            inputDictionary["PortaleDiProvenienza"] = "";

            return inputDictionary;
        }
        // legge il test scenario
        public static List < Dictionary<string, string>> ReadTestScenario (List<string> inputList, string fileNameTestScenario)
        {
            var testList = new List<Dictionary<string, string>>();


        using (var fs = new FileStream(fileNameTestScenario, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0); // Ottieni il primo foglio di lavoro

                // Itera su ogni riga del foglio di lavoro
                //salta la riga 0 con titoli
                for (int row = 1; row <= sheet.LastRowNum; row++)
                {   
                    // Salta le righe vuote
                    if (sheet.GetRow(row) == null)
                        continue;

                    var inputDictionary = new Dictionary<string, string>();
                    // Legge le celle delle colonne di input
                    var col=0;  
                    foreach (var input in inputList)
                    {   
                        ICell cell = sheet.GetRow(row).GetCell(col);
                        if (cell != null)
                        {
                            inputDictionary[input] = cell.ToString();
                        }
                        col++;

                    }
                    testList.Add(inputDictionary);
                }
            }

            return testList;

        }

        // esegue il test scenario
        public static void ExecuteDmnTestScenario(Boolean debug, string fileName, string filePathNameTestScenario)
        {
            try
            {
                var ctx = MyDmn.CreateDmnCtx(fileName);
                var inputList = MyDmn.GetDmnInputs(false,fileName);
                var decision = MyDmn.GetDmnDecision(false,fileName);
                var outputList = MyDmn.GetDmnOutputs(false, fileName);
                var result = new DmnDecisionResult();

                var testInputTable = MyDmn.ReadTestScenario(inputList,filePathNameTestScenario);
                var testOutputTable = new List <Dictionary<string, string>> (testInputTable);
                // init output table
                foreach (var testOutput in testOutputTable)
                {
                    foreach (var output in outputList)
                    {
                        testOutput[output] = "";
                    }
                }

                var testCase=0;
                foreach (var testInput in testInputTable)
                    {
                            MyDmn.LogMessage(debug,"testInput = ");
                            foreach (var input in testInput)
                            {
                                if (inputList.Contains(input.Key))
                                {
                                    MyDmn.LogMessage(debug,"\t" + input.Key + " = " + input.Value);
                                    ctx.WithInputParameter(input.Key, input.Value);
                                }

                            }
                            MyDmn.LogMessage(debug,"eseguo Dmn");
                            MyDmn.LogMessage(debug,"Risultati =");
                            
                            result = ctx.ExecuteDecision(decision);


                            foreach (var item in result.Results)
                            { 
                                foreach (var resVariable in item.Variables)
                                {
                                    if (resVariable.Value != null)
                                        {   
                                            var singleResult = resVariable.Value.ToString();
                                            var allResults = testOutputTable[testCase][resVariable.Name] +"#" + resVariable.Value.ToString();
                                            testOutputTable[testCase][resVariable.Name] = allResults;
                                            }
                                } 
                            } 

                    testCase++;             
                    }
                MyDmn.WriteTestScenario(debug,fileName,filePathNameTestScenario,testOutputTable);
            }
            catch (Exception e)
            {
                throw new DmnException("executeDmnTestScenario error: " + e.Message, e);
            }
        }

        // scrive il test scenario
        public static void WriteTestScenario (Boolean debug, string fileName, string filePathNameTestScenario, List<Dictionary<string, string>> testOutputTable)
        {
            try
            {
                var inputList = MyDmn.GetDmnInputs(false,fileName);
                var outputList = MyDmn.GetDmnOutputs(false, fileName);

                var workbook = new XSSFWorkbook();
                IFont fontRed = workbook.CreateFont();
                fontRed.Color = IndexedColors.Red.Index;
                IFont fontBlk = workbook.CreateFont();
                fontBlk.Color = IndexedColors.Black.Index;

                ICellStyle styleRed = workbook.CreateCellStyle();
                ICellStyle styleBlk = workbook.CreateCellStyle();
                
                styleRed.SetFont(fontRed);
                styleBlk.SetFont(fontBlk);
                

                var sheet = workbook.CreateSheet("DMN Test Scenario");
                var row = 0;
                var cell = 0;
                var headerRow = sheet.CreateRow(row);
                foreach (var input in inputList)
                {
                    headerRow.CreateCell(cell).SetCellValue(input);
                    cell++;
                }
                foreach (var output in outputList)
                {
                    headerRow.CreateCell(cell).SetCellValue(output);
                    cell++;
                }
                row++;

                foreach (var testOutput in testOutputTable)
                {
                    var dataRow = sheet.CreateRow(row);
                    // imposto i valori degli input
                    foreach (var elem in testOutput)
                    {
                        int index = inputList.IndexOf(elem.Key);
                        if (index >= 0)
                        {
                            ICell myCell = dataRow.CreateCell(index);
                            myCell.SetCellValue(elem.Value);
                            myCell.CellStyle = styleBlk;

                        }
                    }
                    // imposto i valori degli output
                    var offset= inputList.Count;
                    foreach (var elem in testOutput)
                    {
                        int index = outputList.IndexOf(elem.Key);
                        if (index >= 0)
                        {
                            ICell myCell = dataRow.CreateCell(offset + index);
                            myCell.SetCellValue(elem.Value);
                            myCell.CellStyle = styleRed;

                        }
                    }
                    row++;
                }
                string executionFileName = fileName.Replace(".dmn", "_test_exec_" + DateTime.Now.ToString("yyyyMMddHHmmss") +".xlsx");
                using (var fileData = new FileStream(executionFileName, FileMode.Create))
                {
                    workbook.Write(fileData);
                }
            }
            catch (Exception e)
            {
                throw new DmnException("WriteTestScenario error: " + e.Message, e);
            }
        }



        // esegue test generico
        public static void ExecuteDmnGenericTest(Boolean debug, string fileName)
        {
            try
            {
                var ctx = MyDmn.CreateDmnCtx(fileName);
                var inputList = MyDmn.GetDmnInputs(false,fileName);
                var decision = MyDmn.GetDmnDecision(false,fileName);
                var outputList = MyDmn.GetDmnOutputs(false, fileName);
                var result = new DmnDecisionResult();

                var testInput = MyDmn.CreateTestInput(inputList);
                foreach (var input in testInput)
                    {
                        MyDmn.LogMessage(debug,"\t" +input.Key + " = " + input.Value);
                        ctx.WithInputParameter(input.Key, input.Value);
                    }
                MyDmn.LogMessage(debug,"eseguo Dmn");
                MyDmn.LogMessage(debug,"Risultati");
                result = ctx.ExecuteDecision(decision);
                foreach (var item in result.Results)
                    { 
                       foreach (var resVariable in item.Variables)
                       {
                           MyDmn.LogMessage(debug, resVariable.ToString());
                        }
                    }
            }        
            catch (Exception e)
            {
                throw new DmnException("executeDmn error: " + e.Message, e);
            }
        }    

    

        public static List<List<string>> GetDmnRules (Boolean debug, string fileName)
        {
            var ruleTable = new List<List<string>>();
            var ruleList = new List<string>();
            var inputList = new List<string>();
            var outputList = new List<string>();

            try
            {
                var doc = XDocument.Load(fileName);
                XNamespace ns = "https://www.omg.org/spec/DMN/20191111/MODEL/";

                var decisionTable = doc.Descendants(ns + "decisionTable").FirstOrDefault();
                if (decisionTable != null)
                {
                    var rules = decisionTable.Descendants(ns + "rule");
                    foreach (var rule in rules)
                    {
                    inputList.Clear();
                    outputList.Clear();
                    foreach (var input in rule.Descendants(ns + "inputEntry"))
                    {
                        var inputText = input.Descendants(ns + "text").FirstOrDefault();
                        inputList.Add(inputText.Value);
                    }

                    foreach (var output in rule.Descendants(ns + "outputEntry"))
                    {
                        var outputText = output.Descendants(ns + "text").FirstOrDefault();
                        outputList.Add(outputText.Value);
                    }
                    ruleList = inputList.Concat(outputList).ToList();
                    ruleTable.Add(ruleList);
                    }
                }

 //               foreach (var item in ruleTable)
 //                   {   
 //                       var ruleString = "";
 //                       foreach (var rule in item)
 //                       {
 //                          ruleString = ruleString + rule + ";";
 //                       }
 //                       MyDmn.LogMessage(debug, ruleString);
 //                   }                
            return ruleTable;
            }
            catch (Exception e)
            {
                throw new DmnException("GetDmnRules error: " + e.Message, e);
            }
        }

        public static void SaveDmnXlsx(Boolean debug, string fileName)
        {
            try
            {

                var ruleTable = new List<List<string>>();
                var inputList = MyDmn.GetDmnInputs(false,fileName);
                var outputList = MyDmn.GetDmnOutputs(false, fileName);

                ruleTable = MyDmn.GetDmnRules(debug, fileName);

                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("DMN Rules");
                var row = 0;
                var cell = 0;
                var headerRow = sheet.CreateRow(row);
                foreach (var input in inputList)
                {
                    headerRow.CreateCell(cell).SetCellValue(input);
                    cell++;
                }
                foreach (var output in outputList)
                {
                    headerRow.CreateCell(cell).SetCellValue(output);
                    cell++;
                }
                row++;
                foreach (var rule in ruleTable)
                {
                    cell = 0;
                    var dataRow = sheet.CreateRow(row);
                    foreach (var item in rule)
                    {
                        dataRow.CreateCell(cell).SetCellValue(item);
                        cell++;
                    }
                    row++;
                }
                string ruleFileName = fileName.Replace(".dmn", "_rules.xlsx");
                using (var fileData = new FileStream(ruleFileName, FileMode.Create))
                {
                    workbook.Write(fileData);
                }
            }
            catch (Exception e)
            {
                throw new DmnException("SaveDmnXlsx error: " + e.Message, e);
            }
        }

        // gest dmn model , for test
        public static void GetDmnModel (Boolean debug, String fileName)
        {
            try
            {
                var model = DmnParser.Parse13ext(fileName);
                var def = DmnDefinitionFactory.CreateDmnDefinition(model);
                var ctx = DmnExecutionContextFactory.CreateExecutionContext(model);
                MyDmn.LogMessage(debug,"model = " + ctx.ToString());
            }        
            catch (Exception e)
            {
                throw new DmnException("executeDmn error: " + e.Message, e);
            }
        }    

    } // end myMyDmn

           // new stage
    public class DmnObject {
            public string dmnFileName { get; set; }
            public Dictionary<string, string> inputDefs { get; set; }
            public Dictionary<string, string> outputDefs { get; set; }

            public List<Dictionary<string, string>> ruleTable { get; set; }
            public string decision { get; set; }

            public void CreateDmnObject (Boolean debug, string fileName)
                {
                inputDefs = new Dictionary<string, string>();
                outputDefs = new Dictionary<string, string>();
                ruleTable = []  ;
                decision = ""   ;

                try
                {
                dmnFileName = fileName;
                var xdoc = XDocument.Load(fileName);
                var ns = xdoc.Root.Name.Namespace;
                var xdecision = xdoc.Descendants(ns + "decision").FirstOrDefault();
                decision = xdecision.Attribute("name").Value;
    
                // colonne di input e loro tipo
                var xinputDefs = xdoc.Descendants(ns + "input");
                foreach (var inputDef in xinputDefs)
                {
                    var label = inputDef.Attribute("label").Value;
                    inputDefs.Add(label, "");
                }

                var xinputExpr = xdoc.Descendants(ns + "inputExpression");
                foreach (var inputExpr in xinputExpr)
                {
                    var typeRef = inputExpr.Attribute("typeRef").Value;
                    var name = inputExpr.Descendants(ns + "text").FirstOrDefault().Value;
                    inputDefs[name] = typeRef;
                }


                // colonne di output di tipo string
                var xoutputDefs = xdoc.Descendants(ns + "output");
                foreach (var outputDef in xoutputDefs)
                {
                    outputDefs.Add(outputDef.Attribute("label").Value, "string");
                }
    
                // regole
                var xrules = xdoc.Descendants(ns + "rule");
                foreach (var rule in xrules)
                {
                    var ruleList = new List<string>();
                    var ruleId = "";
                    var inputEntries = rule.Descendants(ns + "inputEntry");
                    var outputEntries = rule.Descendants(ns + "outputEntry");

                    foreach (var inputEntry in inputEntries)
                    {
                        var inputText = inputEntry.Descendants(ns + "text").FirstOrDefault().Value;
                        ruleList.Add(inputText);
                    }

                    foreach (var outputEntry in outputEntries)
                    {
                        var outputText = outputEntry.Descendants(ns + "text").FirstOrDefault().Value;
                        ruleList.Add(outputText);
                    }


                    var ruleRow = new Dictionary<string, string>();
                    ruleId = rule.Attribute("id").Value;
                    ruleRow.Add("Id", ruleId);

                    var i=0;
                    foreach (var inputDef in inputDefs)
                    {
                        ruleRow.Add(inputDef.Key, ruleList[i]);
                        i++;
                    }
                    foreach (var outputDef in outputDefs)
                    {
                        ruleRow.Add(outputDef.Key, ruleList[i]);
                        i++;
                    }
                    
                    ruleTable.Add(ruleRow);
                }

                }
                catch (Exception e)
                {
                    throw new Exception("createDmn error: " + e.Message, e);
                }
            }
        // add empty rule

        public string AddRule (Boolean debug)
                {
                var ruleId ="";
                try
                {
                    if (dmnFileName == "")
                    {
                        throw new Exception("dmn not initialized");
                    }

                    ruleId = "Rule_" + Guid.NewGuid().ToString().Substring(0, 7);

                    var blankRow = new Dictionary<string, string>();
                    var masterRow = ruleTable[0];

                    foreach (var key in masterRow.Keys.ToList())
                    {
                        blankRow.Add(key,"-");
                    }
                    blankRow["Id"] = ruleId;
                    ruleTable.Add(blankRow);

                }
                catch (Exception e)
                {
                    throw new Exception("add row error: " + e.Message, e);
                }
                return ruleId;
            }
        // delete rule

        public void DeleteRule (Boolean debug, string ruleId)
                {
                try
                {
                    if (dmnFileName == "")
                    {
                        throw new Exception("dmn not initialized");
                    }

                    var rule = ruleTable.Find(x => x["Id"] == ruleId);
                    if (rule != null)
                    {
                        ruleTable.Remove(rule);
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("delete row error: " + e.Message, e);
                }
            }
        // update rule

        public void UpdateRule (Boolean debug, string ruleId, string key, string value)
                {
                try
                {
                    if (dmnFileName == "")
                    {
                        throw new Exception("dmn not initialized");
                    }

                    var rule = ruleTable.Find(x => x["Id"] == ruleId);
                    if (rule != null)
                    {
                        rule[key] = value;
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("update row error: " + e.Message, e);
                }
            } 
    
        // save dmn object to file
        public void SaveDmnObject (Boolean debug ) 
        {
            try
            {
                if (dmnFileName == "")
                {
                    throw new Exception("dmn not initialized");
                }

                var xdoc = XDocument.Load(dmnFileName);
                var ns = xdoc.Root.Name.Namespace;

                // elimino le regole
                var xrules = xdoc.Descendants(ns + "rule");
                foreach (var rule in xrules.ToList())
                {
                    rule.Remove();
                }

                // aggiungo le nuove regole
                foreach (var ruleRow in ruleTable)
                {
                    var xrule = new XElement(ns + "rule");
                    xrule.SetAttributeValue("id", ruleRow["Id"]);
                    var i=0;
                    foreach (var inputDef in inputDefs)
                    {
                        var xinputEntry = new XElement(ns + "inputEntry");
                        var xinputText = new XElement(ns + "text");
                        xinputText.Value = ruleRow[inputDef.Key];
                        xinputEntry.Add(xinputText);
                        xrule.Add(xinputEntry);
                        i++;
                    }
                    foreach (var outputDef in outputDefs)
                    {
                        var xoutputEntry = new XElement(ns + "outputEntry");
                        var xoutputText = new XElement(ns + "text");
                        xoutputText.Value = ruleRow[outputDef.Key];
                        xoutputEntry.Add(xoutputText);
                        xrule.Add(xoutputEntry);
                        i++;
                    }
                    xdoc.Descendants(ns + "decisionTable").FirstOrDefault().Add(xrule);
                }

                xdoc.Save(dmnFileName);
            }
            catch (Exception e)
            {
                throw new Exception("saveDmn error: " + e.Message, e);
            }
        }
    }
    // end new stage
}
