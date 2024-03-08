using net.adamec.lib.common.dmn.engine.parser;
using net.adamec.lib.common.dmn.engine.engine.execution.context;
//using System.Collections.Generic;
//using net.adamec.lib.common.dmn.engine.engine.definition.builder;
//using System.IO;
 
namespace DMN_Engine
{
    class Program
    {
        void Main(string[] args)
        {

            var fileName = @".\diagram_3.dmn";
            var model = DmnParser.Parse13ext(fileName);
            //var def = DmnDefinitionFactory.CreateDmnDefinition(model);
            var ctx = DmnExecutionContextFactory.CreateExecutionContext(model);
            //ctx.WithInputParameter("Season", "Fall");
            //var dyna = new { age = 7};
            //ctx.WithInputParameter("dyna", dyna);



            Console.WriteLine("Sono qui 2!");

            ctx.WithInputParameter("age", 10);
            ctx.WithInputParameter("nazion", "Italia");
            var result = ctx.ExecuteDecision("Decision1");
            Console.WriteLine(result);
            Console.WriteLine(result.First["output_us"]);

            // Definire gli input come dizionario chiave-valore
            //Dictionary<string, object> inputs = new Dictionary<string, object>();
            //inputs.Add("Season", "Fall");
            //inputs.Add("Vegetarian Guests", "");

        }
    }
}
