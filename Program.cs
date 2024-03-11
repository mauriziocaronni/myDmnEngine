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
           }
            catch (Exception e)
                {
                    MyDmn.LogMessage(debug,"error: " + e.Message);
                }
            MyDmn.LogMessage(debug,"end date " + DateTime.Now.ToString()) ;    
        }
    }
}
