using net.adamec.lib.common.dmn.engine.parser;
using net.adamec.lib.common.dmn.engine.engine.execution.context;
using net.adamec.lib.common.dmn.engine.engine.definition;
using net.adamec.lib.common.dmn.engine.engine.execution.result;
using System.Xml.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

 

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
                //for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                for (int row = 0; row <= sheet.LastRowNum; row++)
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
                foreach (var testInput in testInputTable)
                    {
                            MyDmn.LogMessage(debug,"testInput = ");
                            foreach (var input in testInput)
                            {
                                MyDmn.LogMessage(debug,"\t" + input.Key + " = " + input.Value);
                                ctx.WithInputParameter(input.Key, input.Value);
                            }
                            MyDmn.LogMessage(debug,"eseguo Dmn");
                            MyDmn.LogMessage(debug,"Risultati =");
                            
                            result = ctx.ExecuteDecision(decision);
                            foreach (var item in result.Results)
                            { 
                                foreach (var resVariable in item.Variables)
                                {
                                    if (resVariable.Value != null)
                                        {MyDmn.LogMessage(debug, resVariable.Value.ToString());}
                                }              
                            } 

                        }
            }
            catch (Exception e)
            {
                throw new DmnException("executeDmnTestScenario error: " + e.Message, e);
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
    }        
}