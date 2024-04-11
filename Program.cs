using System.Data;
using net.adamec.lib.common.dmn.engine.engine.execution.context;
using net.adamec.lib.common.dmn.engine.engine.execution.result;

namespace MyDmnEngine
{
    class Program
    {

        // main
        static void Main(string[] args)
        {
            // path dei file dmn                        
            var pathDmnTest = @"./dmn/test";
            var pathDmnPreprod = @"./dmn/preprod";

            var debug = true;
            var noDebug = false;
            var execTest = false;

            // inizio
            MyDmn.LogMessage(debug, "start date " + DateTime.Now.ToString());

            try
            {
                // Ottieni tutti i file nella directory
                string[] fileEntries = Directory.GetFiles(pathDmnTest, "*.dmn");
                // Itera su ogni file
                foreach (string fileName in fileEntries)
                {
                    MyDmn.LogMessage(debug,"processing " + fileName);
                    MyDmn.CheckDmnSyntax(debug,fileName);

                    MyDmn.SaveDmnXlsx(debug,fileName);

                    if (execTest)
                    {
                        // verifico presenza testScenario
                        string filePathNameTestScenario = pathDmnTest +"/" + Path.GetFileNameWithoutExtension(fileName) + "_test.xlsx";
                        if (File.Exists(filePathNameTestScenario))
                            {
                            MyDmn.LogMessage(debug,"testScenario found");
                            MyDmn.ExecuteDmnTestScenario(debug,fileName,filePathNameTestScenario);

                            }   
                        else
                            {
                            MyDmn.LogMessage(debug,"testScenario not found");
                            MyDmn.LogMessage(debug,"imposto dati input generici  ");
                            MyDmn.ExecuteDmnGenericTest(debug,fileName);
                            }
                    }

                    // creo un oggetto MyDmnObject da file
                    
                    DmnObject myDmnObject = new DmnObject();
                    myDmnObject.CreateDmnObject(debug,fileName);
                    
                    MyDmn.LogMessage(debug,"DMnObject: " + myDmnObject.ToString());

                    // creo una riga vuota
                    myDmnObject.AddRule(debug);

                    MyDmn.LogMessage(debug,"NewDMnObject: " + myDmnObject.ToString());

                    // modifico un valore 
                    myDmnObject.UpdateRule(debug, "Rule_0kug95k", "Azienda", "\"Beta80\"");
                    MyDmn.LogMessage(debug,"NewDMnObject: " + myDmnObject.ToString());

                    // cancello una riga
                    myDmnObject.DeleteRule(debug, "Rule_1u35u6s");
                    MyDmn.LogMessage(debug,"NewDMnObject: " + myDmnObject.ToString());

                    // rilsalvo il file
                    myDmnObject.SaveDmnObject(debug);

                }        
           }
            catch (Exception e)
                {
                    MyDmn.LogMessage(debug,"error: " + e.Message);
                }
            MyDmn.LogMessage(debug,"end date " + DateTime.Now.ToString()) ;    
        }
    }
}
